import csv
import os
import re
from typing import List, Dict, Tuple

MASTER_HEADER = ["id", "page", "x", "y", "w", "h", "zoom", "method", "nominal", "lsl", "usl", "bx", "by", "br"]


def atomic_write(path: str, rows: List[Dict[str, str]], header: List[str]):
    tmp = path + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=header)
        writer.writeheader()
        for r in rows:
            writer.writerow({k: ("" if r.get(k) is None else r.get(k)) for k in header})
    os.replace(tmp, path)


def ensure_master(pdf_path: str) -> str:
    master = pdf_path + ".balloons.csv"
    if not os.path.exists(master):
        atomic_write(master, [], MASTER_HEADER)
    return master


def read_master(pdf_path: str) -> List[Dict[str, str]]:
    master = ensure_master(pdf_path)
    rows = []
    with open(master, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append(r)
    return rows


def write_master(pdf_path: str, rows: List[Dict[str, str]]):
    master = ensure_master(pdf_path)
    atomic_write(master, rows, MASTER_HEADER)


def next_id(rows: List[Dict[str, str]]) -> str:
    maxn = 0
    for r in rows:
        vid = r.get("id", "")
        m = re.search(r"(\d+)$", vid)
        if m:
            n = int(m.group(1))
            if n > maxn:
                maxn = n
    return f"HS-{maxn+1:03d}"


def add_feature(pdf_path: str, feature: Dict[str, str]) -> Dict[str, str]:
    rows = read_master(pdf_path)
    fid = next_id(rows)
    feature_row = {k: "" for k in MASTER_HEADER}
    feature_row.update(feature)
    feature_row["id"] = fid
    # ensure bx/by exist
    feature_row.setdefault("bx", "0")
    feature_row.setdefault("by", "0")
    feature_row.setdefault("br", "14")
    rows.append(feature_row)
    write_master(pdf_path, rows)
    return feature_row


def update_feature(pdf_path: str, fid: str, updates: Dict[str, str]):
    rows = read_master(pdf_path)
    changed = False
    for r in rows:
        if r.get("id") == fid:
            r.update(updates)
            changed = True
            break
    if changed:
        write_master(pdf_path, rows)


def delete_feature(pdf_path: str, fid: str) -> bool:
    rows = read_master(pdf_path)
    new_rows = [r for r in rows if r.get("id") != fid]
    if len(new_rows) == len(rows):
        return False
    write_master(pdf_path, new_rows)
    return True


# Work Order CSV helpers
def wo_path(pdf_path: str, workorder: str) -> str:
    safe = workorder.replace("/", "_")
    return f"{pdf_path}.{safe}.csv"


def read_wo(pdf_path: str, workorder: str) -> Dict[str, str]:
    path = wo_path(pdf_path, workorder)
    if not os.path.exists(path):
        return {}
    res = {}
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            res[r.get("id")] = r.get("result")
    return res


def write_wo(pdf_path: str, workorder: str, data: Dict[str, str]):
    path = wo_path(pdf_path, workorder)
    rows = []
    for id_, result in data.items():
        rows.append({"id": id_, "result": result})
    atomic_write(path, rows, ["id", "result"])


def list_workorders(pdf_path: str) -> List[str]:
    base_name = os.path.basename(pdf_path)
    directory = os.path.dirname(pdf_path) or "."
    prefix = f"{base_name}."
    results: List[str] = []
    try:
        entries = os.listdir(directory)
    except FileNotFoundError:
        return results
    for name in entries:
        if not name.startswith(prefix) or not name.endswith(".csv"):
            continue
        if name == f"{base_name}.balloons.csv":
            continue
        full_path = os.path.join(directory, name)
        if not os.path.isfile(full_path):
            continue
        token = name[len(prefix):-4]
        if not token:
            continue
        results.append(token.replace("_", "/"))
    results.sort(key=lambda s: s.lower())
    return results


# Normalization for parsing tolerances
def normalize_str(s: str) -> str:
    if s is None:
        return ""
    # replace various unicode minus/plus with ascii
    s = s.replace("−", "-").replace("＋", "+").replace("－", "-")
    s = s.replace("Ø", "").replace("ø", "")
    return s.strip()


def parse_tolerance_expression(s: str) -> Tuple[float, float, float]:
    """Parse strings like:
    12.34 ±0.1
    12.34 +0.1 -0.2
    12.34
    Returns (nominal, lsl, usl) or raises ValueError
    """
    s2 = normalize_str(s)
    if not s2:
        raise ValueError("empty")

    # equal bilateral: look for ± or +/- or \u00B1
    m = re.fullmatch(r"([+-]?\d*\.?\d+)\s*(?:±|\+/-|/\-\+|\+\-|\+/?-?)\s*([+-]?\d*\.?\d+)", s2)
    if m:
        nom = float(m.group(1))
        tol = abs(float(m.group(2)))
        return nom, nom - tol, nom + tol

    # alternative: explicit +/- with spaces: nominal +a -b  (order may vary)
    m = re.fullmatch(r"([+-]?\d*\.?\d+)\s*([+-]\s*\d*\.?\d+)\s*([+-]\s*\d*\.?\d+)", s2)
    if m:
        nom = float(m.group(1))
        a = float(m.group(2).replace(" ", ""))
        b = float(m.group(3).replace(" ", ""))
        # find which is plus and which is minus
        plus = None
        minus = None
        if a >= 0 and b <= 0:
            plus = a; minus = -b
        elif b >= 0 and a <= 0:
            plus = b; minus = -a
        else:
            # sign included in the numbers, compute lsl/usl directly
            lsl = nom + min(a, b)
            usl = nom + max(a, b)
            return nom, lsl, usl
        return nom, nom - minus, nom + plus

    # plain number
    m = re.fullmatch(r"([+-]?\d*\.?\d+)", s2)
    if m:
        nom = float(m.group(1))
        # decimals
        text = m.group(1)
        if '.' in text:
            dec = len(text.split('.')[-1])
        else:
            dec = 0
        if dec == 1:
            tol = 0.03
        elif dec == 2:
            tol = 0.01
        else:
            tol = 0.005
        return nom, nom - tol, nom + tol

    raise ValueError("Unrecognized format")
