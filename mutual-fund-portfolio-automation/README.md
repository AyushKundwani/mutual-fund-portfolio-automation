#### **Project: Mutual Fund Portfolio Data Consolidation \& Automation**

Python-based data pipeline to consolidate and standardize mutual fund portfolio disclosures into analytics-ready datasets.



###### **1. Data Model and Assumptions**



* **Data Model**



 	The portfolio holdings from multiple mutual fund scheme sheets are consolidated into a structured tabular format.

 	Each row in the output dataset represents one instrument holding within a scheme.



* **Final schema used:**



| **Column Name**       | **Description**                                        |

| ----------------- | -------------------------------------------------- |

| amc\_name          | Name of Asset Management Company                   |

| scheme\_name       | Scheme name derived from sheet name                |

| instrument\_name   | Name of the security/instrument                    |

| instrument\_type   | Asset class classification (Equity / Debt / Other) |

| isin              | ISIN code of the security                          |

| market\_value      | Market/Fair value or amount invested               |

| portfolio\_percent | Percentage weight of the instrument in portfolio   |

| report\_date       | Portfolio reporting date                           |



* **Assumptions:**



1. The Excel workbook contains one sheet per scheme.
2. Each scheme sheet contains a tabular section with portfolio holdings.
3. Header rows may appear after introductory text, so header row detection is dynamic.
4. ISIN is treated as a reliable identifier for valid instrument rows.
5. Footer sections such as "Scheme Risk-O-Meter" are not part of the portfolio table and are removed.
6. Instrument type classification is rule-based using keywords in instrument names.
7. Portfolio percentage values may appear in different formats and are standardized during processing.







**2. Automation Approach**



* The solution is designed as an automated data ingestion and normalization pipeline:



**A. Workbook Parsing:**



* **All sheets are read programmatically.**
* **Index or non-portfolio sheets are skipped.**



**B. Header Detection:**



* **The script searches for the row containing portfolio column headers dynamically rather than assuming a fixed row.**



**C. Column Standardization:**



* Column names are cleaned (lowercase, symbols removed).
* Market value and portfolio percentage columns are auto-detected using keyword matching.
* Columns are mapped into a standardized schema.



**D. Data Cleaning:**



* Non-data rows and footer sections are removed.
* Only rows with valid ISIN values are retained.
* Missing placeholders such as "-", "NA", or empty strings are converted to nulls.



**E. Metadata Enrichment:**



* AMC name is inferred from workbook content.
* Scheme name is taken from the sheet name.
* Reporting date is added as a metadata field.



**F. Instrument Classification:**



* Instruments are categorized into Equity, Debt, or Other based on name-based keyword rules.



G. **Output Segmentation:**



* **The consolidated dataset is split into:**

      1. Equity portfolio file

      2. Debt portfolio file



###### **3. Steps to Run the Code**



* **The generated CSV files will be saved in the same directory where the Python script is located.**



* **Install required libraries:**



 	1. pip install pandas openpyxl numpy



* **Update the Excel file path in the script:**



 	2. file\_path = "your\_portfolio\_file.xlsx"





* **Run the Python script.**



* **Output files will be generated:**



 	1.equity\_portfolio.csv

 	2.debt\_portfolio.csv


