from excel_extract import Extract
from excel_transform import Transform


# Step 1: Extract data
extractor = Extract("test_file_nc.xlsx")
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
else:
    print("Failed to load data. Check file path and format.")
