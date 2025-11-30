from __future__ import annotations

import math
import os
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional

from storage import read_master, list_workorders, read_wo, wo_path


@dataclass
class Measurement:
    feature_id: str
    value: float
    workorder: str
    source_path: str
    timestamp: Optional[datetime] = None


@dataclass
class FeatureStats:
    count: int = 0
    mean: Optional[float] = None
    stdev: Optional[float] = None
    min_value: Optional[float] = None
    max_value: Optional[float] = None
    cp: Optional[float] = None
    cpk: Optional[float] = None


@dataclass
class FeatureSPCData:
    feature_id: str
    method: str = ""
    nominal: Optional[float] = None
    lsl: Optional[float] = None
    usl: Optional[float] = None
    measurements: List[Measurement] = field(default_factory=list)
    stats: FeatureStats = field(default_factory=FeatureStats)


def _parse_float(value: str) -> Optional[float]:
    if value is None:
        return None
    text = value.strip()
    if not text:
        return None
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def _parse_spec(value: str) -> Optional[float]:
    if not value:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _compute_stats(values: List[float], lsl: Optional[float], usl: Optional[float]) -> FeatureStats:
    stats = FeatureStats()
    if not values:
        return stats
    stats.count = len(values)
    stats.min_value = min(values)
    stats.max_value = max(values)
    stats.mean = sum(values) / stats.count
    if stats.count > 1:
        variance = sum((v - stats.mean) ** 2 for v in values) / (stats.count - 1)
        stats.stdev = math.sqrt(max(variance, 0.0))
    else:
        stats.stdev = 0.0
    if lsl is not None and usl is not None and stats.stdev and stats.stdev > 0:
        tol = usl - lsl
        stats.cp = tol / (6.0 * stats.stdev)
        upper_term = (usl - stats.mean) / (3.0 * stats.stdev)
        lower_term = (stats.mean - lsl) / (3.0 * stats.stdev)
        stats.cpk = min(upper_term, lower_term)
    return stats


def load_spc_dataset(pdf_path: str) -> Dict[str, FeatureSPCData]:
    master_rows = read_master(pdf_path)
    feature_map: Dict[str, FeatureSPCData] = {}
    for row in master_rows:
        fid = row.get("id", "").strip()
        if not fid:
            continue
        feature_map[fid] = FeatureSPCData(
            feature_id=fid,
            method=(row.get("method") or "").strip(),
            nominal=_parse_spec(row.get("nominal")),
            lsl=_parse_spec(row.get("lsl")),
            usl=_parse_spec(row.get("usl")),
        )

    directory = os.path.dirname(pdf_path) or "."
    wo_entries = list_workorders(pdf_path)
    for wo in wo_entries:
        csv_path = wo_path(pdf_path, wo)
        timestamp = None
        try:
            timestamp = datetime.fromtimestamp(os.path.getmtime(csv_path))
        except OSError:
            pass
        results = read_wo(pdf_path, wo)
        for fid, text in results.items():
            value = _parse_float(text)
            if value is None:
                continue
            feature = feature_map.get(fid)
            if not feature:
                continue
            if os.path.exists(csv_path):
                try:
                    relative_path = os.path.relpath(csv_path, directory)
                except ValueError:
                    relative_path = csv_path
            else:
                relative_path = csv_path
            feature.measurements.append(
                Measurement(
                    feature_id=fid,
                    value=value,
                    workorder=wo,
                    source_path=relative_path,
                    timestamp=timestamp,
                )
            )

    for feature in feature_map.values():
        if not feature.measurements:
            continue
        measurement_values = [m.value for m in feature.measurements]
        feature.stats = _compute_stats(measurement_values, feature.lsl, feature.usl)

    # Return only features with data
    return {fid: data for fid, data in feature_map.items() if data.measurements}


def format_stat(value: Optional[float], decimals: int = 3) -> str:
    if value is None:
        return "—"
    return f"{value:.{decimals}f}"