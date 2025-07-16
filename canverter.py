import pdfplumber
import pandas as pd
import numpy as np

pdf_path = "data2.pdf"
all_tables = []

with pdfplumber.open(pdf_path) as pdf:
    # go through each page of pdf
    for page_num, page in enumerate(pdf.pages, start=1):
        # extract table from pdf
        tables = page.extract_tables()

        for table_index, table in enumerate(tables, start=1):
            # Convert table to DataFrame
            df = pd.DataFrame(table)

            # clean dataframe by empty rows and columns
            df.dropna(how="all", inplace=True)
            
            # Convert DataFrame to NumPy array
            np_data = df.to_numpy()

            # count page(where table has) or numbers of table 
            print(f"Page {page_num}: {len(tables)} tables found")
            # Print NumPy array shape
            print("NumPy array shape:", np_data.shape)

            # error handling for numeric operations
            try:
                # convert dataframe to numericc data 
                numeric_data = df.apply(pd.to_numeric, errors='coerce')
                # calculate column wise means and ignoreing nan value
                means = np.nanmean(numeric_data.to_numpy(), axis=0)
                print("Column-wise means:", means)
            except:
                # handle nan data 
                print("Non-numeric data or mixed types in table")

            all_tables.append(df)

# save all table in excel sheet
with pd.ExcelWriter("output.xlsx", engine="openpyxl") as writer:
    for i, table_df in enumerate(all_tables):
    # check if the dataframe is empty or not
      if not table_df.empty:
        # create a defferent-2 sheet for table
          sheet_name = f"Table_{i+1}"
          print(f"Writing table to sheet: {sheet_name}, shape: {table_df.shape}")
        # write table to excel sheet
          table_df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False, header=False)

print(f"Total tables found in PDF: {len(all_tables)}")