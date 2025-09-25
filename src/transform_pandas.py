# transform_pandas.py
from excel_extractor import ExcelExtractor
import pandas as pd
from typing import Dict, Optional, List
import re


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

print(df_samples)



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

class SampleReporter:
    def __init__(self) -> None:
        pass

'''
qc = QCReporter(df)
print(qc.table())
'''







'''

# for calibration, will figure this part in the future. IN THE WORKS
# will have version

wanted_for_calibration = [r for r in records if r.sample_type == "Update Standard"]

df_for_calibration = pd.DataFrame([r.__dict__ for r in wanted_for_calibration])

print(df_for_calibration)

'''
