import pandas as pd
import openpyxl

df = pd.read_csv(r"C:\Users\123ca\Documents\Scraped Data\Superstore.csv", encoding="latin1")

# Numerical rounding for sales data
df["Profit"] = df["Profit"].round(2)
df["Sales"] = df["Sales"].round(2)

# Converting the date fields to standardized datetime formats for Power BI use
df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")

# Standardizing date formats. YYYY-MM-DD works best for Power BI
df["Order Date"] = df["Order Date"].dt.date
df["Ship Date"] = df["Ship Date"].dt.date

# Putting the data frame data into an excel sheet.
df.to_excel(r"C:\Users\123ca\Documents\Scraped Data\Superstore_Fixed.xlsx", index=False)
print("Succesfully downloaded!")