import pdfplumber
import pandas as pd
import re


def extract_hdfc_statement(pdf_path, password=None):
    """
    Extracts HDFC bank statement tables from a multi-page PDF into a Pandas DataFrame.
    """
    all_transactions = []

    # Open the PDF (support password if needed)
    with pdfplumber.open(pdf_path, password=password) as pdf:
        print(f"Processing {len(pdf.pages)} pages in '{pdf_path}'...")

        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()

            for table in tables:
                for row in table:
                    clean_row = [cell.strip() if cell else "" for cell in row]
                    if len(clean_row) < 7:
                        continue

                    all_transactions.append(clean_row)

    if not all_transactions:
        print(
            "Error: No data found. The PDF might be an image scan or have a different layout."
        )
        return None

    # Create Initial DataFrame
    df = pd.DataFrame(all_transactions)

    headers = [
        "Date",
        "Narration",
        "Chq./Ref.No.",
        "Value Date",
        "Withdrawal Amount",
        "Deposit Amount",
        "Closing Balance",
    ]

    header_mask = (df[0].str.contains("Date", case=False, na=False)) & (
        df[1].str.contains("Narration", case=False, na=False)
    )

    if header_mask.any():
        df.columns = headers
        df = df[~header_mask]
    else:
        df.columns = df.iloc[0]
        df = df[1:]

    cols_to_clean = ["Withdrawal Amount", "Deposit Amount", "Closing Balance"]
    for col in cols_to_clean:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(",", "", regex=False)
            df[col] = df[col].astype(str).str.replace(r"[^\d.]", "", regex=True)
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # 2. Clean & Format Dates
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
        # Step B: Format back to String (DD/MM/YYYY) to match your visual requirement
        df["Date"] = df["Date"].dt.strftime("%d/%m/%Y")

    if "Value Date" in df.columns:
        df["Value Date"] = pd.to_datetime(
            df["Value Date"], dayfirst=True, errors="coerce"
        )
        df["Value Date"] = df["Value Date"].dt.strftime("%d/%m/%Y")

    # 3. Clean Narration
    # PDF tables often wrap text with newlines (\n). We replace them with spaces.
    if "Narration" in df.columns:
        df["Narration"] = (
            df["Narration"].astype(str).str.replace("\n", " ", regex=False)
        )

    df = df.reset_index(drop=True)

    return df


if __name__ == "__main__":
    # 1. SETUP: Put your exact PDF filename here
    pdf_filename = "statement.pdf"  # <--- CHANGE THIS to your file name

    # 2. RUN EXTRACTION
    try:
        # If your PDF has a password, add it here: extract_hdfc_statement(pdf_filename, password="123")
        df_result = extract_hdfc_statement(pdf_filename)

        if df_result is not None:
            # 3. PREVIEW
            print("\n--- Preview of Extracted Data ---")
            print(df_result.head(10))

            # 4. SAVE TO EXCEL/CSV
            output_file = "hdfc_cleaned_data.xlsx"
            df_result.to_excel(output_file, index=False)
            print(f"\nSuccess! Data saved to: {output_file}")

    except FileNotFoundError:
        print(f"Error: The file '{pdf_filename}' was not found. Please check the name.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
