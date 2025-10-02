# transform_pandas.py
from excel_extractor import ExcelExtractor
import pandas as pd
from typing import Dict, Optional, List

class Format:


    #for sample class
    @staticmethod
    def group_pairs(df: pd.DataFrame, group_col: str = "sample_id"):
        pairs = []
        for _, group in df.groupby(group_col):
            group = group.sort_index()
            pair_list = [group.iloc[i:i+2] for i in range(0, len(group), 2)]
            pairs.extend(pair_list)
        return pairs


ex = ExcelExtractor("test_file_nc.xlsx", sheet="", header_row_index=1)
records = list(ex.records())


# ----- helpers -----
def norm(s) -> str:
    # Uppercase, strip edges, collapse inner spaces (e.g., "ccv 1" -> "CCV1")
    if s is None:
        return ""
    return str(s).strip().upper().replace(" ", "")

def is_rinse(sample_id: str) -> bool:
    return "RINSE" in norm(sample_id)

def is_qc(sample_id: str) -> bool:
    """QC if MDL/QCS/QCB or numbered CCV#/CCB# (case/space insensitive)."""
    sid = norm(sample_id)
    if sid in {"MDL", "QCS", "QCB"}:
        return True
    if sid.startswith("CCV") and sid[3:].isdigit():
        return True
    if sid.startswith("CCB") and sid[3:].isdigit():
        return True
    return False

# QC records (by sample_id pattern), regardless of sample_type
wanted_QC = [r for r in records if is_qc(r.sample_id)]

# Samples = sample_type == "Samples" AND not rinses AND not QC-looking IDs
wanted_Samples = [
    r for r in records
    if str(r.sample_type).strip() == "Samples"
    and not is_rinse(r.sample_id)
    and not is_qc(r.sample_id)
]

# DataFrames
df_qc      = pd.DataFrame([r.__dict__ for r in wanted_QC])
df_samples = pd.DataFrame([r.__dict__ for r in wanted_Samples])

# QC REPORT
class QCReporter:
    def __init__(self, df: pd.DataFrame, targets: Optional[Dict[str, float]] = None):
        cols = [c for c in ["sample_id", "ppm"] if c in df.columns]
        self.df = df[cols].copy()

        # Default targets
        default_targets = {"MDL": 0.2, "QCS": 10.0, "CCV": 10.0}
        self.targets = default_targets.copy()
        if targets:
            self.targets.update({k: float(v) for k, v in targets.items()})

    def grab(self) -> pd.DataFrame:
        pattern = r"^(MDL|QCS|QCB|CCB\d+|CCV\d+)$"
        mask = self.df["sample_id"].astype(str).str.match(pattern, na=False)
        return self.df[mask].copy()

    def _percent_r(self, avg: float, target: Optional[float]) -> Optional[float]:
        return (avg / target) * 100.0 if (avg is not None and target) else None

    def _rpd(self, values: List[float]) -> Optional[float]:
        if values is None or len(values) < 2:
            return None
        v1, v2 = float(values[0]), float(values[1])
        mean = (v1 + v2) / 2.0
        if mean == 0:
            return None
        return abs(v1 - v2) / mean * 100.0

    def table(self) -> pd.DataFrame:
        qc = self.grab().copy()
        qc["sample_id"] = qc["sample_id"].astype(str)

        output_rows = []

        # ---- Handle MDL ----
        mdl_rows = qc[qc["sample_id"] == "MDL"]
        if not mdl_rows.empty:
            output_rows.extend(mdl_rows.to_dict("records"))
            vals = mdl_rows["ppm"].tolist()
            avg = sum(vals) / len(vals)
            output_rows.append({
                "sample_id": "Average_MDL",
                "ppm": avg,
                "%R": self._percent_r(avg, self.targets["MDL"]),
                "%RPD": self._rpd(vals)
            })

        # ---- Handle QCS ----
        qcs_rows = qc[qc["sample_id"] == "QCS"]
        if not qcs_rows.empty:
            output_rows.extend(qcs_rows.to_dict("records"))
            vals = qcs_rows["ppm"].tolist()
            avg = sum(vals) / len(vals)
            output_rows.append({
                "sample_id": "Average_QCS",
                "ppm": avg,
                "%R": self._percent_r(avg, self.targets["QCS"]),
                "%RPD": self._rpd(vals)
            })

        # ---- Handle each CCV# separately ----
        ccv_rows = qc[qc["sample_id"].str.match(r"^CCV\d+$")]
        if not ccv_rows.empty:
            ccv_ids = sorted(ccv_rows["sample_id"].unique(),
                             key=lambda s: int(s[3:]))
            for cid in ccv_ids:
                block = ccv_rows[ccv_rows["sample_id"] == cid]
                output_rows.extend(block.to_dict("records"))
                vals = block["ppm"].tolist()
                avg = sum(vals) / len(vals)
                output_rows.append({
                    "sample_id": f"Average_{cid}",
                    "ppm": avg,
                    "%R": self._percent_r(avg, self.targets["CCV"]),
                    "%RPD": self._rpd(vals)
                })

        # ---- Handle QCB + CCB together ----
        qcbccb_rows = qc[qc["sample_id"].str.match(r"^(QCB|CCB\d+)$")]
        if not qcbccb_rows.empty:
            # order: QCB, then CCB1, CCB2, ...
            qcb_rows = qcbccb_rows[qcbccb_rows["sample_id"] == "QCB"]
            ccb_rows = qcbccb_rows[qcbccb_rows["sample_id"].str.match(r"^CCB\d+$")]

            if not qcb_rows.empty:
                output_rows.extend(qcb_rows.to_dict("records"))

            if not ccb_rows.empty:
                ccb_ids = sorted(ccb_rows["sample_id"].unique(),
                                 key=lambda s: int(s[3:]))
                for cid in ccb_ids:
                    block = ccb_rows[ccb_rows["sample_id"] == cid]
                    output_rows.extend(block.to_dict("records"))

            # combined average
            vals = qcbccb_rows["ppm"].tolist()
            avg = sum(vals) / len(vals)
            output_rows.append({
                "sample_id": "Average_QCB_CCB",
                "ppm": avg,
                "%R": None,
                "%RPD": None
            })

        return pd.DataFrame(output_rows)


# ---------------- Calculations (integrated) ----------------
class Calculations:
    @staticmethod
    def avg(values):
        if values is None:
            return None
        vals = [float(v) for v in values if v is not None]
        return sum(vals) / len(vals) if vals else None

    @staticmethod
    def rpd(v1, v2):
        if v1 is None or v2 is None:
            return None
        v1, v2 = float(v1), float(v2)
        mean_val = (v1 + v2) / 2.0
        if mean_val == 0:
            return None
        return abs(v1 - v2) / mean_val * 100.0

    @staticmethod
    def pair_avg_df(df_pair, value_col="mean"):
        if df_pair is None or value_col not in df_pair.columns:
            return None
        return Calculations.avg(df_pair[value_col].tolist())

    @staticmethod
    def pair_rpd_df(df_pair, value_col="mean"):
        if df_pair is None or len(df_pair) != 2 or value_col not in df_pair.columns:
            return None
        v1, v2 = df_pair[value_col].tolist()
        return Calculations.rpd(v1, v2)
    
    @staticmethod
    def to_umol_per_l_c(avg_mean, factor: float = 83.26):
        """Convert an average mean to µmol/L C using the given factor."""
        if avg_mean is None:
            return None
        return float(avg_mean) * float(factor)


class SampleReporter:
    def __init__(self, value_col: str = "mean", factor: float = 83.26) -> None:
        self.value_col = value_col
        self.factor = factor

    def compute_pair_average(self, df_pair: pd.DataFrame, value_col: str | None = None) -> float:
        vcol = value_col or self.value_col
        return Calculations.pair_avg_df(df_pair, value_col=vcol)

    def compute_pair_rpd(self, df_pair: pd.DataFrame, value_col: str | None = None) -> float | None:
        vcol = value_col or self.value_col
        return Calculations.pair_rpd_df(df_pair, value_col=vcol)

    def compute_pair_umol_per_l_c(self, df_pair: pd.DataFrame, value_col: str | None = None, factor: float | None = None) -> float:
        vcol = value_col or self.value_col
        f = factor if factor is not None else self.factor
        avg_mean = self.compute_pair_average(df_pair, value_col=vcol)
        return Calculations.to_umol_per_l_c(avg_mean, factor=f)



'''
# Group pairs automatically
sr = SampleReporter()
pairs = Format.group_pairs(df_samples)



print(df_samples)

# Optional: see the grouped pairs + calculations
for i, p in enumerate(pairs, 1):
    sid = str(p.iloc[0]["sample_id"]) if not p.empty else "NA"
    avg_mean = sr.compute_pair_average(p, value_col="mean")
    rpd_mean = sr.compute_pair_rpd(p, value_col="mean")
    print(f"\n--- Pair {i} (sample_id={sid}) ---")
    print(p)
    print(f"Average(mean): {avg_mean}")
    print(f"%RPD(mean): {rpd_mean}")


for p in pairs:
    sid = str(p.iloc[0]["sample_id"]) if not p.empty else "NA"
    avg_mean = sr.compute_pair_average(p, "mean")
    umol = sr.compute_pair_umol_per_l_c(p, "mean", factor=83.26)
    print(f"{sid} → avg(mean)={avg_mean}, µmol/L C={umol}")
'''