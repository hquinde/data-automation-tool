import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "code"))

from excel_extract import Extract
from excel_transform import Transform
from excel_load import Load

# Find Excel file in input_files
input_dir = Path("input_files")
excel_files = list(input_dir.glob("*.xlsx"))

if excel_files:
    file_path = excel_files[0]
    print(f"Processing: {file_path.name}")
    
    extractor = Extract(str(file_path))
    raw_data = extractor.extract_data()
    
    if raw_data is not None:  # Changed this line
        transformer = Transform(raw_data)
        loader = Load(transformer)
        loader.export_all()
        print("Done! Check output_files/results.xlsx")
else:
    print("No Excel file found in input_files/")