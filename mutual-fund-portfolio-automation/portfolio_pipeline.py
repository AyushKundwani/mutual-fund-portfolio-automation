import pandas as pd
import numpy as np
import os

# file path

file_path = "D:\Project\Monthly Portfolio-31 12 25.xlsx"

# Loading of workbook

xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names

print("Sheets found:", sheet_names)

# Creating dataframe

master_df = pd.DataFrame()

# Detect AMC name from first sheet

first_sheet_df = pd.read_excel(file_path, sheet_name=0, header=None)
text_block = " ".join(first_sheet_df.head(10).astype(str).values.flatten()).lower()

if "axis" in text_block:
    amc_name = "Axis Mutual Fund"
elif "sbi" in text_block:
    amc_name = "SBI Mutual Fund"
elif "hdfc" in text_block:
    amc_name = "HDFC Mutual Fund"
else:
    amc_name = "Unknown AMC"

report_date = "2025-12-31"

# Looping through all schemes

for sheet in sheet_names:

    if sheet.lower() == "index":
        continue

    print(f"\nProcessing sheet: {sheet}")

    df = pd.read_excel(file_path, sheet_name=sheet, header=None)

    header_row = None
    for i, row in df.iterrows():
        if row.astype(str).str.contains("Name of the Instrument", case=False).any():
            header_row = i
            break

    if header_row is None:
        print(f" Header not found in {sheet}, skipping...")
        continue

    # Setting the proper header

    df.columns = df.iloc[header_row]
    df = df[header_row + 1:]

    # Here the Cleaning part begins

    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace("\n", " ")
        .str.replace(" ", "_")
        .str.replace("%", "percent")
        .str.replace("/", "_")
        .str.replace("(", "")
        .str.replace(")", "")
    )

    # Detecting the market value columns

    market_col = None
    for col in df.columns:
        if any(word in col for word in ["market", "fair", "value", "amount"]):
            market_col = col
            break

    # Detecting the Portfolio % Column

    percent_col = None
    for col in df.columns:
        if any(word in col for word in ["percent", "nav", "net_asset"]):
            percent_col = col
            break

    if percent_col:
        df.rename(columns={percent_col: "portfolio_percent"}, inplace=True)
        print(f" Portfolio % column detected: {percent_col}")
    else:
        print(f" Portfolio % column not found in {sheet}")

    for col in df.columns:
        col_check = col.lower()
        if "name_of_the_instrument" in col_check:
            df.rename(columns={col: "instrument_name"}, inplace=True)

        if "market" in col_check and "value" in col_check:
            df.rename(columns={col: "market_value"}, inplace=True)

        if "percent" in col_check and "net" in col_check:
            df.rename(columns={col: "portfolio_percent"}, inplace=True)

    #Removing the schema risk o meter as it not the part of the portfolio

    risk_index = df[
        df.astype(str)
          .apply(lambda row: row.str.contains("risk-o-meter", case=False, na=False))
          .any(axis=1)
    ].index

    if len(risk_index) > 0:
        df = df.loc[:risk_index[0]-1]

    if "isin" in df.columns:
        df = df[df["isin"].notna()]
    else:
        print(f" ISIN column missing in {sheet}, skipping...")
        continue

    df["scheme_name"] = sheet
    df["amc_name"] = amc_name
    df["report_date"] = report_date

    # Detecting of Instrument Type

    def detect_type(name):
        if isinstance(name, str):
            name = name.lower()
            debt_keywords = [
                "government", "govt", "treasury", "t-bill", "bill",
                "debenture", "ncd", "bond", "sdl", "certificate of deposit",
                "commercial paper", "cp"
            ]
            equity_keywords = [
                "ltd", "limited", "industries", "corp", "corporation",
                "technologies", "bank", "finance"
            ]

            if any(word in name for word in debt_keywords):
                return "Debt"
            elif any(word in name for word in equity_keywords):
                return "Equity"
            else:
                return "Other"
        return "Other"

    df["instrument_type"] = df["instrument_name"].apply(detect_type)

    # Keeping Only the required Columns

    final_cols = [
        "amc_name",
        "scheme_name",
        "instrument_name",
        "instrument_type",
        "isin",
        "market_value",
        "portfolio_percent",
        "report_date"
    ]

    df = df[[col for col in final_cols if col in df.columns]]

    # Concating to master_df

    master_df = pd.concat([master_df, df], ignore_index=True)

# Final Cleaning

master_df.replace(["-", "NA", ""], np.nan, inplace=True)
master_df = master_df.infer_objects(copy=False)

# Converting the percentage from decimal to percentage

if "portfolio_percent" in master_df.columns:
    master_df["portfolio_percent"] = (
        master_df["portfolio_percent"]
        .astype(str)
        .str.replace("%", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
    )

    master_df["portfolio_percent"] = (
        pd.to_numeric(master_df["portfolio_percent"], errors="coerce") * 100
    ).round(2).astype(str) + "%"

equity_df = master_df[master_df["instrument_type"] == "Equity"]
debt_df = master_df[master_df["instrument_type"] == "Debt"]

# In this step we are splitting it based on the instrument type

equity_df.to_csv("equity_portfolio.csv", index=False)
debt_df.to_csv("debt_portfolio.csv", index=False)


# Exporting it to CSV file in which equity goes to equity_portfolio and debt goes to debt_portfolio
print("Files saved in current directory:", os.getcwd())
print("File successfully Exported")
print(" Equity records:", len(equity_df))
print(" Debt records:", len(debt_df))
