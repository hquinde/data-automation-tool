from excel_extract import Extract

# Example use
ex = Extract("test_file.xlsx", header_row_index = 1)
for rec in ex.records():
    print(rec.sample_id, rec.sample_type, rec.mean, rec.ppm, rec.adjusted_abs)




