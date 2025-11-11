import pandas as pd


class Transform:
    def __init__(self, df):
        self.df = df


    def clean_data(self):
        cleaned = self.df[self.df["Sample Type"] == "Samples"]
        print(f"Cleaned data: {len(cleaned)} samples remaining")
        return cleaned


    def group_samples(self, df, group_col="Sample ID"):
        grouped_dfs = []
        for _, group in df.groupby(group_col, sort=False):
            group_reset = group.reset_index(drop=True)
            grouped_dfs.append(group_reset)
        return grouped_dfs


    def filter_qcb_ccb(self, df):
        sample_id = df["Sample ID"].astype("string").str.strip()
        mask = sample_id.str.fullmatch(r"ICB|CCB\d*", na=False) # changed QCB to ICB
        filtered = df[mask]
        return filtered


    def calculate_mean_ppm(self, df):
        return df["PPM"].mean()


    def calculate_rpd(self, df, mean_ppm):
        values = df["PPM"].dropna().tolist()
        v1 = float(values[-2])
        v2 = float(values[-1])
        rpd = abs(v1 - v2) / mean_ppm * 100.0
        return rpd


    def calculate_percent_R(self, df, target_override=None):
        values = df["PPM"].dropna()
        mean_value = float(values.mean())
        
        if target_override is not None:
            target = float(target_override)
        else:
            targets = df["Mean (per analysis type)"].dropna()
            target = float(targets.iloc[0])
        
        percent_r = mean_value / target * 100.0
        return percent_r


    def convert_to_umol_per_L(self, ppm_value, molecular_weight):
        # Formula: ppm * 1000 / molecular_weight 
        # For carbon (12.01057): 1000/12.01057 ≈ 83.26 (used in Excel)
        return float(ppm_value) * 1000.0 / molecular_weight