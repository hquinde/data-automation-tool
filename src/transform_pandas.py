# transform_pandas.py
from excel_extractor import ExcelExtractor
import pandas as pd

ex = ExcelExtractor("test_file.xlsx", sheet="Data", header_row_index=1)
records = list(ex.records())

wanted = {} # edit this so that it gets sample type != Blank or Clean

sample_ids   = [r.sample_id for r in records]
means        = [r.mean for r in records]
ppms         = [r.ppm for r in records]
adjusted_abs = [r.adjusted_abs for r in records]

df = pd.DataFrame({
    "Sample ID": sample_ids,
    "Mean": means,
    "PPM": ppms,
    "Adjusted ABS": adjusted_abs
})
print(df.head())


# note for calibration, we need update standard.... think about this.

