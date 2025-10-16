from typing import List, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from excel_extract import Extract
from excel_transform import Transform


class Load:
    def __init__(self, transformer, output_path: str = "results.xlsx", molecular_weight: float = 12.01057):
        if transformer is None:
            raise ValueError("Load requires a Transform instance")
        self.transformer = transformer
        self.output_path = output_path
        self.molecular_weight = molecular_weight

    @staticmethod
    def is_out_of_bounds(value, check_type):
        """
        Check if value is out of bounds.
        Returns True if OUT OF BOUNDS (make it red).
        """
        try:
            val = float(value)
        except:
            return True
        
        if check_type == 'QC_R':
            return val < 90 or val > 110
        elif check_type == 'MDL_R':
            return val < 45 or val > 145
        elif check_type == 'RPD':
            return val > 10
        else:
            return True

    def _sample_groups(self) -> Tuple[pd.DataFrame, List[Tuple[str, pd.DataFrame]]]:
        cleaned = self.transformer.clean_data()
        if cleaned.empty:
            return cleaned, []

        df = cleaned.copy()
        df["Sample ID"] = df["Sample ID"].astype("string").str.strip()

        qc_pattern = r"(?i)^(MDL|QCS|QCB|CCV\d+|CCB\d+|MQ\s*(?:Auto)?Rinse)$"
        samples_only = df[~df["Sample ID"].str.match(qc_pattern, na=False)]
        if samples_only.empty:
            return samples_only, []

        ordered_ids: List[str] = []
        seen = set()
        for sid in samples_only["Sample ID"]:
            if sid not in seen:
                seen.add(sid)
                ordered_ids.append(sid)

        groups: List[Tuple[str, pd.DataFrame]] = []
        for sample_id in ordered_ids:
            group_df = samples_only[samples_only["Sample ID"] == sample_id].copy()
            if not group_df.empty:
                groups.append((sample_id, group_df))

        return samples_only, groups
    
    def format_calibration(self):
        df = self.transformer.df
        if "Sample Type" not in df.columns:
            return pd.DataFrame()

        sample_type = df["Sample Type"].astype("string")
        mask = sample_type.str.contains("cal", case=False, na=False)
        calibration = df[mask].copy()

        if "Sample ID" in calibration.columns:
            calibration["Sample ID"] = calibration["Sample ID"].astype("string")
            calibration.sort_values(by=["Sample ID"], inplace=True, ignore_index=True)
        return calibration
    
    def format_qc(self):
        df = self.transformer.df
        if df is None or df.empty or "Sample ID" not in df.columns:
            return pd.DataFrame(columns=["Sample ID", "PPM C", "Mean ppm C", "%R", "%RPD"])

        samples = df[df.get("Sample Type") == "Samples"].copy()
        if samples.empty:
            return pd.DataFrame(columns=["Sample ID", "PPM C", "Mean ppm C", "%R", "%RPD"])

        samples["Sample ID"] = samples["Sample ID"].astype("string").str.strip()

        qc_mask = samples["Sample ID"].str.match(r"(?i)^(MDL|QCS|CCV\d*)$", na=False)
        qcb_mask = samples["Sample ID"].str.match(r"(?i)^(QCB|CCB\d*)$", na=False)

        qc_samples = samples[qc_mask]
        qcb_samples = samples[qcb_mask]

        columns = ["Sample ID", "PPM C", "Mean ppm C", "%R", "%RPD"]
        records = []

        qc_targets = {
            "MDL": 0.2,
            "QCS": 18.0,
            "CCV1": 10.0,
            "CCV2": 10.0,
        }

        if not qc_samples.empty:
            for sample_id in qc_samples["Sample ID"].unique():
                group_df = qc_samples[qc_samples["Sample ID"] == sample_id]
                if group_df.empty:
                    continue

                group_records = []
                for _, row in group_df.iterrows():
                    group_records.append(
                        {
                            "Sample ID": row["Sample ID"],
                            "PPM C": row.get("PPM"),
                            "Mean ppm C": None,
                            "%R": None,
                            "%RPD": None,
                        }
                    )

                mean_ppm = self.transformer.calculate_mean_ppm(group_df)
                target = qc_targets.get(str(sample_id).upper())
                percent_r = self.transformer.calculate_percent_R(group_df, nominal_override=target)
                rpd = self.transformer.calculate_rpd(group_df, mean_ppm)

                if group_records:
                    last_record = group_records[-1]
                    if mean_ppm is not None:
                        last_record["Mean ppm C"] = mean_ppm
                    if percent_r is not None:
                        last_record["%R"] = percent_r
                    if rpd is not None:
                        last_record["%RPD"] = rpd

                records.extend(group_records)

        if records and not qcb_samples.empty:
            records.append({col: None for col in columns})

        if not qcb_samples.empty:
            for _, row in qcb_samples.iterrows():
                records.append(
                    {
                        "Sample ID": row["Sample ID"],
                        "PPM C": row.get("PPM"),
                        "Mean ppm C": None,
                        "%R": None,
                        "%RPD": None,
                    }
                )

            average_ppm = self.transformer.calculate_mean_ppm(qcb_samples)
            records.append(
                {
                    "Sample ID": "Average",
                    "PPM C": average_ppm if average_ppm is not None else None,
                    "Mean ppm C": None,
                    "%R": None,
                    "%RPD": None,
                }
            )

        return pd.DataFrame(records, columns=columns)
    
    def format_samples(self):
        samples_only, groups = self._sample_groups()
        if samples_only.empty or not groups:
            return pd.DataFrame(columns=["Sample ID", "PPM C", "Mean ppm C", "%RPD", "umol/L C"])

        columns = ["Sample ID", "PPM C", "Mean ppm C", "%RPD", "umol/L C"]
        records = []

        for sample_id, group_df in groups:
            group_records = []
            for _, row in group_df.iterrows():
                ppm_value = row.get("PPM")
                umol_value = self.transformer.convert_to_umol_per_L(ppm_value, self.molecular_weight)
                group_records.append(
                    {
                        "Sample ID": row["Sample ID"],
                        "PPM C": row.get("PPM"),
                        "Mean ppm C": None,
                        "%RPD": None,
                        "umol/L C": umol_value,
                    }
                )

            mean_ppm = self.transformer.calculate_mean_ppm(group_df)
            rpd = self.transformer.calculate_rpd(group_df, mean_ppm)
            mean_umol = self.transformer.convert_to_umol_per_L(mean_ppm, self.molecular_weight)

            if group_records:
                last_record = group_records[-1]
                if mean_ppm is not None:
                    last_record["Mean ppm C"] = mean_ppm
                if rpd is not None:
                    last_record["%RPD"] = rpd
                if mean_umol is not None:
                    last_record["umol/L C"] = mean_umol

            records.extend(group_records)

        return pd.DataFrame(records, columns=columns)
    
    def format_reported_results(self):
        _, groups = self._sample_groups()
        if not groups:
            return pd.DataFrame(columns=["Sample ID", "umol/L C"])

        records = []
        for sample_id, group_df in groups:
            mean_ppm = self.transformer.calculate_mean_ppm(group_df)
            umol = self.transformer.convert_to_umol_per_L(mean_ppm, self.molecular_weight)

            records.append(
                {
                    "Sample ID": sample_id,
                    "umol/L C": umol if umol is not None else None,
                }
            )

        return pd.DataFrame(records, columns=["Sample ID", "umol/L C"])
    
    def export_all(self):
        sheets = {
            "Calibration": (self.format_calibration(), {}),
            "QC": (self.format_qc(), {}),
            "Samples": (self.format_samples(), {}),
            "Reported Results": (self.format_reported_results(), {}),
        }

        wrote_any = False
        try:
            # Write data with openpyxl engine
            with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                for sheet_name, (df, options) in sheets.items():
                    if df is None or df.empty:
                        print(f"Skipping sheet '{sheet_name}': no rows to export")
                        continue
                    export_options = {"index": False}
                    if options:
                        export_options.update(options)
                    df.to_excel(writer, sheet_name=sheet_name, **export_options)
                    wrote_any = True
                    print(f"Wrote {len(df)} rows to sheet '{sheet_name}'")
            
            # Apply red text formatting
            if wrote_any:
                wb = load_workbook(self.output_path)
                red_font = Font(color='FF0000', bold=True)
                
                # Format QC sheet - check %R column
                if 'QC' in wb.sheetnames:
                    ws = wb['QC']
                    for row_idx in range(2, ws.max_row + 1):
                        # Get Sample ID from column A (1)
                        sample_id_cell = ws.cell(row=row_idx, column=1)
                        sample_id = sample_id_cell.value
                        
                        # %R is in column D (4)
                        r_cell = ws.cell(row=row_idx, column=4)
                        if r_cell.value is not None:
                            # Determine check type based on Sample ID
                            if sample_id and 'MDL' in str(sample_id).upper():
                                check_type = 'MDL_R'
                            else:
                                check_type = 'QC_R'
                            
                            if self.is_out_of_bounds(r_cell.value, check_type):
                                r_cell.font = red_font
                
                # Format Samples sheet - check %RPD column
                if 'Samples' in wb.sheetnames:
                    ws = wb['Samples']
                    for row_idx in range(2, ws.max_row + 1):
                        # %RPD is in column D (4)
                        rpd_cell = ws.cell(row=row_idx, column=4)
                        if rpd_cell.value is not None:
                            if self.is_out_of_bounds(rpd_cell.value, 'RPD'):
                                rpd_cell.font = red_font
                
                wb.save(self.output_path)
                print(f"Applied bounds checking and formatting")
                
        except ImportError as exc:
            print(f"Failed to export Excel file: {exc}")
            return

        if wrote_any:
            print(f"Export finished: {self.output_path}")
        else:
            print("No data exported — all sheets were empty.")


# Step 1: Extract data
extractor = Extract("TEST2.xlsx")
raw_data = extractor.extract_data()

if raw_data is not None:
    # Step 2: Initialize transformer
    transformer = Transform(raw_data)
    
    # Step 3: Clean data (filter to samples only)
    cleaned_df = transformer.clean_data()
    
    # Step 4: Group the cleaned data
    grouped_samples = transformer.group_samples(cleaned_df)
    print(f"\nFirst group preview:")
    print(grouped_samples[0] if grouped_samples else "No groups")
    
    # Step 5: Filter QCB/CCB from the cleaned data
    qcb_ccb_df = transformer.filter_qcb_ccb(cleaned_df)
    print(f"\nQCB/CCB Data:")
    print(qcb_ccb_df)
    
    # Step 6: Calculate QCB/CCB average
    qcb_ccb_average = transformer.calculate_mean_ppm(qcb_ccb_df)
    print(f"\nQCB/CCB Average: {qcb_ccb_average}")
    
    # Step 7: Calculate Mean and RPD for each sample group
    print("\n--- Sample Group Calculations ---")
    for group_df in grouped_samples:
        sample_id = group_df["Sample ID"].iloc[0] if not group_df.empty else "Unknown"
        
        # Calculate mean for the group
        group_mean = transformer.calculate_mean_ppm(group_df)
        
        # Calculate RPD (passing the mean as denominator)
        rpd = transformer.calculate_rpd(group_df, group_mean)
        
        print(f"{sample_id}: Mean={group_mean}, %RPD={rpd}")

    loader = Load(transformer)
    loader.export_all()
else:
    print("Failed to load data. Check file path and format.")
