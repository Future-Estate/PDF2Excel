import pdfplumber
import pandas as pd

pdf_path = "clients.pdf"
excel_path = "clients.xlsx"

all_tables = []

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        tables = page.extract_tables()
        for table in tables:
            df = pd.DataFrame(table[1:], columns=table[0])
            # Reset index to avoid duplicate index issues
            df = df.reset_index(drop=True)
            # Remove duplicate columns, keep first occurrence
            df = df.loc[:, ~df.columns.duplicated()]
            all_tables.append(df)

if all_tables:
    final_df = pd.concat(all_tables, ignore_index=True)
    final_df.to_excel(excel_path, index=False)
    print(f"Data successfully written to {excel_path}")
else:
    print("No tables found in the PDF.")
