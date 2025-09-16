# excel_extractor.py
from openpyxl import load_workbook
from typing import Iterator, Optional
from extract_base import Extractor, Record

class ExcelExtractor(Extractor):

    def __init__(self, path: str, sheet: str | None = None, header_row_index: int = 1):
        """
        path: path to the Excel file
        sheet: specific sheet name, or None to use the first sheet -> be ready to edit this later for other machines
        header_row_index: which row has the headers (1-based, usually 1)
        """
        self.path = path
        self.sheet = sheet
        self.header_row_index = header_row_index


    @staticmethod
    def _to_float(value) -> float | None: # changes value in cell into a float

        if value is None or value == "": #might have to edit this because idk how empty well would be
            return None
        try:
            return float(value)
        except (TypeError) or (ValueError):
            return None

    @staticmethod
    def _to_str(value) -> str | None:
        return "" if value is None else str(value).strip()
    

    def _pick_sheet(self, wb):
        # If the user specified a sheet, honor it
        if self.sheet:
            return wb[self.sheet]

        # Otherwise, scan all sheets and pick the first that has our headers
        wanted = {"Sample ID", "Mean (per analysis type)", "PPM", "Adjusted ABS"}
        for name in wb.sheetnames:
            ws = wb[name]
            rows = ws.iter_rows(values_only=True)
            for _ in range(self.header_row_index - 1):
                next(rows, None)
            header = list(next(rows, []))
            normalized = { (str(h).strip() if h is not None else "") for h in header }
            if any(h in normalized for h in wanted):
                return ws
        # Fallback: active sheet
        return wb.active


    def _find_header_map(self, ws) -> dict[str, int]:
        rows = ws.iter_rows(values_only=True)

        # jump to the header row
        for _ in range(self.header_row_index - 1):
            next(rows, None)

        header = list(next(rows, []))  # the header line (row with titles)

        wanted = {"Sample ID", "Mean (per analysis type)", "PPM", "Adjusted ABS"}

        header_to_idx: dict[str, int] = {}
        for idx, name in enumerate(header):
            if isinstance(name, str):
                name_clean = name.strip()
                if name_clean in wanted:
                    header_to_idx[name_clean] = idx

        missing = wanted - set(header_to_idx.keys())
        if missing:
            raise ValueError(f"Missing required header(s): {', '.join(missing)}")

        return header_to_idx


    def records(self) -> Iterator[Record]: # see if you can change it
        wb = load_workbook(self.path, read_only=True, data_only=True)
        try:
            ws = self._pick_sheet(wb)

            # get where each header is
            header_map = self._find_header_map(ws)

            # start reading AFTER the header row
            rows = ws.iter_rows(values_only=True)
            for _ in range(self.header_row_index):  # skip header rows
                next(rows, None)

            for row in rows:
                # helper to fetch a cell by header name safely
                def at(name: str):
                    i = header_map[name]
                    return row[i] if i < len(row) else None

                sample_id    = self._to_str(at("Sample ID"))
                mean         = self._to_float(at("Mean (per analysis type)"))
                ppm          = self._to_float(at("PPM"))
                adjusted_abs = self._to_float(at("Adjusted ABS"))

                # skip totally empty lines
                if not sample_id and mean is None and ppm is None and adjusted_abs is None:
                    continue

                yield Record( sample_id=sample_id, mean=mean, ppm=ppm, adjusted_abs=adjusted_abs)

        finally:
            wb.close()

#----#
# Example use

ex = ExcelExtractor("test_file.xlsx", sheet=None, header_row_index=1)

for rec in ex.records():
    print(rec.sample_id, rec.mean, rec.ppm, rec.adjusted_abs)

#----#





        


    








