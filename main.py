import tabula
import pandas as pd
import re
import time


# Path to your PDF file
pdf_path = "1.pdf"

SUB_INDEX = "סאב\rאינדקס"
ARTICLE_NUM = "מס'\rארטיקל"
PRODUCT_NUM = "מס'\rפרודקט"

tables = tabula.read_pdf(pdf_path, pages="all", multiple_tables=True)

indexes_of_interest = ["F 05", "F 40", "F 70", "F 82"]

df = pd.concat(tables, ignore_index=True)
all_categories = indexes_of_interest + [x for x in df[SUB_INDEX].unique() if x not in indexes_of_interest]
df[SUB_INDEX] = pd.Categorical(df[SUB_INDEX], categories=all_categories, ordered=True)
df = df.sort_values(SUB_INDEX).reset_index(drop=True)

# df = df[df[SUB_INDEX].isin(indexes_of_interest)]
df[ARTICLE_NUM] = df[ARTICLE_NUM].apply(lambda x: re.sub("[^0-9]", "", str(x)))

def build_link(row):
    product_num = str(row[PRODUCT_NUM]).strip()
    article_num = str(row[ARTICLE_NUM]).strip()
    return f"https://www2.hm.com/hw_il/productpage.{product_num}{article_num}.html"


df["link"] = df.apply(build_link, axis=1)

# Save to Excel with hyperlinks
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

wb = Workbook()
ws = wb.active

# Write dataframe to worksheet
for r, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
    for c, value in enumerate(row, 1):
        cell = ws.cell(row=r, column=c, value=value)
        # If this is the "מס'\rפרודקט" column (not header), add hyperlink
        if r > 1 and ws.cell(row=1, column=c).value == "מס'\rפרודקט":
            link_value = df.iloc[r-2]['link']
            cell.hyperlink = link_value
            cell.style = "Hyperlink"  # Optional: style as hyperlink

wb.save('output.xlsx')
print("Excel file 'output.xlsx' created with hyperlinks in the product number column.")