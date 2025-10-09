import pandas as pd
from typing import Optional


class Extract:
    def __init__(self, file_path: str, header_row_index: int = 1):
        self.file_path = file_path
        self.header_row_index = header_row_index

    @staticmethod
    def normalize(s):
        return "" if s is None else str(s).strip()

    def extract_data(self, cols=("Sample ID", "Sample Type", "Mean (per analysis type)", "PPM", "Adjusted ABS")) -> Optional[pd.DataFrame]:
        """Extract data from Excel file with specified columns."""
        want = {self.normalize(c) for c in cols}

        def selector(col_name):
            return self.normalize(col_name) in want

        try:
            df = pd.read_excel(self.file_path, header=self.header_row_index - 1, usecols=selector)
            print(f"Successfully loaded {len(df)} rows from {self.file_path}")
            return df
        except FileNotFoundError:
            print(f"Error: File not found at {self.file_path}")
            return None
        except ValueError as error:
            print(f"Error reading columns: {error}")
            return None