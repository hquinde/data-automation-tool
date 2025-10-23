import pandas as pd

class Extract:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.header_row_index = 1
        self.cols = ("Sample ID", "Sample Type", "Mean (per analysis type)", "PPM", "Adjusted ABS")


    def extract_data(self):
        wanted_columns = set()
        for column in self.cols:
            wanted_columns.add(column.strip())
        
        def should_include_column(column_name):
            if column_name.strip() in wanted_columns:
                return True
            else:
                return False
        


        # print statements for debugging
        try:
            df = pd.read_excel(
                self.file_path, 
                header=self.header_row_index - 1,
                usecols=should_include_column
            )
            
            print(f"Successfully loaded {len(df)} rows from {self.file_path}")
            return df
            
        except FileNotFoundError:
            print(f"Error: File not found at {self.file_path}")
            return None
            
        except ValueError as error:
            print(f"Error reading columns: {error}")
            return None