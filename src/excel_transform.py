import pandas as pd
from typing import List, Optional


class Transform:
    """
    Refactored Transform class with consistent, immutable design.
    All methods return new DataFrames without modifying the original.
    """
    
    def __init__(self, df: pd.DataFrame):
        if df is None or df.empty:
            raise ValueError("Cannot initialize Transform with None or empty DataFrame")
        self.df = df.copy()  # Store a copy to avoid side effects

    def clean_data(self) -> pd.DataFrame:
        """Filter to only 'Samples' type rows."""
        cleaned = self.df[self.df["Sample Type"] == "Samples"].copy()
        print(f"Cleaned data: {len(cleaned)} samples remaining")
        return cleaned

    def group_samples(self, df: pd.DataFrame, group_col: str = "Sample ID") -> List[pd.DataFrame]:
        """
        Groups rows by group_col, preserving original order.
        Returns a list of DataFrames — one per unique group.
        """
        if df.empty:
            print("Warning: No data available to group.")
            return []

        if group_col not in df.columns:
            raise KeyError(f"Column '{group_col}' not found in DataFrame.")

        grouped_dfs = []
        for _, group in df.groupby(group_col, sort=False):
            grouped_dfs.append(group.reset_index(drop=True))
        
        print(f"Grouped into {len(grouped_dfs)} unique samples")
        return grouped_dfs

    def filter_qcb_ccb(self, df: pd.DataFrame) -> pd.DataFrame:
        """Filter to only QCB and CCB samples."""
        sample_id = df["Sample ID"].astype("string").str.strip()
        mask = sample_id.str.fullmatch(r"QCB|CCB\d*", na=False)
        filtered = df[mask].copy()
        print(f"Filtered to {len(filtered)} QCB/CCB samples")
        return filtered

    def calculate_mean_ppm(self, df: pd.DataFrame) -> float:
        """Calculate mean PPM for given dataframe."""
        mean = df["PPM"].mean()
        return mean

    def calculate_rpd(self, df: pd.DataFrame, mean_ppm: float, value_col: str = "PPM") -> Optional[float]:
        """
        Calculate Relative Percent Difference for a group of samples.
        Formula: %RPD = (|run1 - run2| / Mean PPM C) * 100
        Expects at least 2 values. Returns None if insufficient data.
        """
        if df.empty or len(df) < 2:
            return None
        
        values = df[value_col].dropna().tolist()
        if len(values) < 2:
            return None
        
        v1, v2 = float(values[-2]), float(values[-1])
        
        if mean_ppm == 0:
            return None
        
        rpd = abs(v1 - v2) / mean_ppm * 100.0
        return rpd

    def calculate_percent_R(
        self,
        df: pd.DataFrame,
        target_col: str = "Mean (per analysis type)",
        value_col: str = "PPM",
        nominal_override: Optional[float] = None,
    ) -> Optional[float]:
        """
        Calculate Percent Recovery relative to the nominal concentration.
        Uses the first non-null value in `target_col` as the reference.
        """
        if df.empty:
            return None

        if value_col not in df.columns:
            return None

        values = df[value_col].dropna()
        if values.empty:
            return None

        if nominal_override is not None:
            nominal = float(nominal_override)
        else:
            if target_col not in df.columns:
                return None
            targets = df[target_col].dropna()
            if targets.empty:
                return None
            nominal = float(targets.iloc[0])

        if nominal == 0:
            return None

        mean_value = float(values.mean())
        percent_r = mean_value / nominal * 100.0
        return percent_r

    def convert_to_umol_per_L(self, ppm_value: Optional[float], molecular_weight: float) -> Optional[float]:
        """Convert a PPM concentration to µmol/L using the provided molecular weight."""
        if ppm_value is None or molecular_weight == 0:
            return None

        return float(ppm_value) * 1000.0 / molecular_weight
