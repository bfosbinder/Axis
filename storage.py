import csv
import getpass
import os
import re
import sqlite3
from contextlib import closing
from datetime import datetime
from typing import Dict, List, Tuple

MASTER_HEADER = ["id", "page", "x", "y", "w", "h", "zoom", "method", "nominal", "lsl", "usl", "bx", "by", "br", "username"]
DB_SUFFIX = ".axis.db"


def _db_path(pdf_path: str) -> str:
    return f"{pdf_path}{DB_SUFFIX}"


def ensure_master(pdf_path: str) -> str:
    """Ensure the SQLite database exists and is initialized; migrate legacy CSV data when present."""
    db_path = _db_path(pdf_path)
    init_needed = not os.path.exists(db_path)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        if init_needed:
            _initialize_db(conn)
        _ensure_feature_schema(conn)
        _maybe_migrate_from_csv(pdf_path, conn)
    finally:
        conn.close()
    return db_path


def _initialize_db(conn: sqlite3.Connection):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS features (
            id TEXT PRIMARY KEY,
            page TEXT,
            x TEXT,
            y TEXT,
            w TEXT,
            h TEXT,
            zoom TEXT,
            method TEXT,
            nominal TEXT,
            lsl TEXT,
            usl TEXT,
            bx TEXT,
            by TEXT,
            br TEXT,
            username TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS results (
            feature_id TEXT NOT NULL,
            workorder TEXT NOT NULL,
            result TEXT,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (feature_id, workorder),
            FOREIGN KEY (feature_id) REFERENCES features(id) ON DELETE CASCADE
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_results_workorder ON results(workorder)")
    conn.commit()


def _connect(pdf_path: str) -> sqlite3.Connection:
    ensure_master(pdf_path)
    conn = sqlite3.connect(_db_path(pdf_path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _ensure_feature_schema(conn: sqlite3.Connection):
    cursor = conn.execute("PRAGMA table_info(features)")
    columns = {row[1] for row in cursor.fetchall()}
    if "username" not in columns:
        conn.execute("ALTER TABLE features ADD COLUMN username TEXT")
        conn.commit()


def _maybe_migrate_from_csv(pdf_path: str, conn: sqlite3.Connection):
    master_csv = f"{pdf_path}.balloons.csv"
    try:
        count = conn.execute("SELECT COUNT(*) FROM features").fetchone()[0]
    except sqlite3.DatabaseError:
        count = 0
    if os.path.exists(master_csv) and count == 0:
        with open(master_csv, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        for row in rows:
            payload = {key: row.get(key, "") or "" for key in MASTER_HEADER}
            conn.execute(
                f"INSERT OR REPLACE INTO features ({', '.join(MASTER_HEADER)}) VALUES ({', '.join('?' for _ in MASTER_HEADER)})",
                [payload.get(key, "") for key in MASTER_HEADER]
            )
        conn.commit()

    # migrate workorder CSVs only once (when DB is still empty)
    try:
        result_count = conn.execute("SELECT COUNT(*) FROM results").fetchone()[0]
    except sqlite3.DatabaseError:
        result_count = 0
    if result_count == 0:
        directory = os.path.dirname(pdf_path) or "."
        base_name = os.path.basename(pdf_path)
        prefix = f"{base_name}."
        try:
            entries = os.listdir(directory)
        except FileNotFoundError:
            entries = []
        for name in entries:
            if not name.startswith(prefix) or not name.endswith(".csv"):
                continue
            if name == f"{base_name}.balloons.csv":
                continue
            workorder = name[len(prefix):-4].replace("_", "/")
            csv_path = os.path.join(directory, name)
            with open(csv_path, newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
            timestamp = datetime.fromtimestamp(os.path.getmtime(csv_path)).isoformat()
            for row in rows:
                feature_id = row.get("id")
                result = row.get("result", "")
                if not feature_id:
                    continue
                conn.execute(
                    """
                    INSERT INTO results (feature_id, workorder, result, updated_at)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(feature_id, workorder)
                    DO UPDATE SET result=excluded.result, updated_at=excluded.updated_at
                    """,
                    (feature_id, workorder, result, timestamp)
                )
            conn.commit()


def read_master(pdf_path: str) -> List[Dict[str, str]]:
    with closing(_connect(pdf_path)) as conn:
        cur = conn.execute(f"SELECT {', '.join(MASTER_HEADER)} FROM features ORDER BY id COLLATE NOCASE")
        rows = []
        for row in cur.fetchall():
            data = {key: (row[key] if row[key] is not None else "") for key in MASTER_HEADER}
            rows.append(data)
        return rows


def write_master(pdf_path: str, rows: List[Dict[str, str]]):
    with closing(_connect(pdf_path)) as conn:
        conn.execute("BEGIN")
        existing_ids = {r[0] for r in conn.execute("SELECT id FROM features").fetchall()}
        incoming_ids = set()
        for row in rows:
            fid = row.get("id")
            if not fid:
                continue
            incoming_ids.add(fid)
            payload = {key: (row.get(key, "") or "") for key in MASTER_HEADER}
            placeholders = ", ".join("?" for _ in MASTER_HEADER)
            columns = ", ".join(MASTER_HEADER)
            if fid in existing_ids:
                assignments = ", ".join(f"{col} = ?" for col in MASTER_HEADER if col != "id")
                values = [payload[col] for col in MASTER_HEADER if col != "id"] + [fid]
                conn.execute(f"UPDATE features SET {assignments} WHERE id = ?", values)
            else:
                conn.execute(f"INSERT INTO features ({columns}) VALUES ({placeholders})", [payload[key] for key in MASTER_HEADER])
        removed = existing_ids - incoming_ids
        if removed:
            conn.executemany("DELETE FROM features WHERE id = ?", [(fid,) for fid in removed])
        conn.commit()


def _next_feature_id(conn: sqlite3.Connection) -> str:
    maxn = 0
    for row in conn.execute("SELECT id FROM features"):
        vid = row[0] or ""
        m = re.search(r"(\d+)$", vid)
        if m:
            maxn = max(maxn, int(m.group(1)))
    return f"{maxn + 1:03d}"


def add_feature(pdf_path: str, feature: Dict[str, str]) -> Dict[str, str]:
    with closing(_connect(pdf_path)) as conn:
        fid = _next_feature_id(conn)
        payload = {key: "" for key in MASTER_HEADER}
        payload.update(feature)
        payload["id"] = fid
        payload.setdefault("bx", "0")
        payload.setdefault("by", "0")
        payload.setdefault("br", "14")
        conn.execute(
            f"INSERT INTO features ({', '.join(MASTER_HEADER)}) VALUES ({', '.join('?' for _ in MASTER_HEADER)})",
            [payload.get(col, "") for col in MASTER_HEADER]
        )
        conn.commit()
        return payload


def update_feature(pdf_path: str, fid: str, updates: Dict[str, str]):
    if not fid or not updates:
        return
    with closing(_connect(pdf_path)) as conn:
        assignments = []
        values = []
        for key, value in updates.items():
            if key not in MASTER_HEADER or key == "id":
                continue
            assignments.append(f"{key} = ?")
            values.append(value)
        if not assignments:
            return
        values.append(fid)
        conn.execute(f"UPDATE features SET {', '.join(assignments)} WHERE id = ?", values)
        conn.commit()


def delete_feature(pdf_path: str, fid: str) -> bool:
    if not fid:
        return False
    with closing(_connect(pdf_path)) as conn:
        cur = conn.execute("DELETE FROM features WHERE id = ?", (fid,))
        conn.commit()
        return cur.rowcount > 0


def current_username() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return os.environ.get("USERNAME") or os.environ.get("USER") or ""


def wo_path(pdf_path: str, workorder: str) -> str:
    """Return the legacy CSV path if it exists, else the SQLite db file for labeling purposes."""
    safe = workorder.replace("/", "_")
    legacy = f"{pdf_path}.{safe}.csv"
    if os.path.exists(legacy):
        return legacy
    return _db_path(pdf_path)


def read_wo(pdf_path: str, workorder: str) -> Dict[str, str]:
    if not workorder:
        return {}
    with closing(_connect(pdf_path)) as conn:
        cur = conn.execute(
            "SELECT feature_id, result FROM results WHERE workorder = ?",
            (workorder,)
        )
        return {row[0]: (row[1] or "") for row in cur.fetchall()}


def write_wo(pdf_path: str, workorder: str, data: Dict[str, str]):
    if not workorder:
        return
    timestamp = datetime.utcnow().isoformat()
    with closing(_connect(pdf_path)) as conn:
        conn.execute("BEGIN")
        conn.execute("DELETE FROM results WHERE workorder = ?", (workorder,))
        for feature_id, result in data.items():
            if not feature_id:
                continue
            conn.execute(
                "INSERT INTO results (feature_id, workorder, result, updated_at) VALUES (?, ?, ?, ?)",
                (feature_id, workorder, result, timestamp)
            )
        conn.commit()


def get_workorder_timestamp(pdf_path: str, workorder: str):
    if not workorder:
        return None
    with closing(_connect(pdf_path)) as conn:
        row = conn.execute("SELECT MAX(updated_at) FROM results WHERE workorder = ?", (workorder,)).fetchone()
    if row and row[0]:
        try:
            return datetime.fromisoformat(row[0])
        except ValueError:
            return None
    return None


def list_workorders(pdf_path: str) -> List[str]:
    with closing(_connect(pdf_path)) as conn:
        cur = conn.execute("SELECT DISTINCT workorder FROM results ORDER BY LOWER(workorder)")
        return [row[0] for row in cur.fetchall() if row[0]]


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
    m = re.fullmatch(r"([+-]?\d*\.?\d+)\s*([+-]\s*\d*\.?\d+)(?:\s*/\s*)?\s*([+-]\s*\d*\.?\d+)", s2)
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
