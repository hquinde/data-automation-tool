# transform_pandas.py
from excel_extractor import ExcelExtractor
import pandas as pd

ex = ExcelExtractor("test_file_nc.xlsx", sheet="", header_row_index=1)
records = list(ex.records())


# filter out unwanted records first
wanted = [
    r for r in records
    if r.sample_type == "Samples"
    and "rinse" not in str(r.sample_id).lower()
]

# build a dataframe from all remaining record attributes
df = pd.DataFrame([r.__dict__ for r in wanted])

#print(df)

# QC REPORT

class QCReporter:

    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()

    def grab(self) -> pd.DataFrame:

        mask = self.df["sample_id"].str.match(r"^(QCB|CCB\d+)$", na=False)
        return self.df[mask].copy()
    
    def calculate_average(self) -> float:
        grabbed = self.grab()
        return grabbed["ppm"].mean()
    
    def table(self) -> pd.DataFrame:
        # grab only sample_id + ppm
        t = self.grab()[["sample_id", "ppm"]].copy()
        avg_ppm = self.calculate_average()

        # make a DataFrame for the average row
        avg_row = {"sample_id": "Average PPM", "ppm": avg_ppm}
        avg_df = pd.DataFrame([avg_row])

        # append using pd.concat (append is removed in pandas 2.0+)
        out = pd.concat([t, avg_df], ignore_index=True)

        return out






qc = QCReporter(df)
#qcb_ccb = qc.grab()
#qcb_av = qc.calculate_average()
#print(qcb_ccb)
#print("Average PPM: ", qcb_av)


print(qc.table())







'''

# for calibration, will figure this part in the future. IN THE WORKS
# will have version

wanted_for_calibration = [r for r in records if r.sample_type == "Update Standard"]

df_for_calibration = pd.DataFrame([r.__dict__ for r in wanted_for_calibration])

print(df_for_calibration)

'''
