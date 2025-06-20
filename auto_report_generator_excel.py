import pandas as pd
import glob
import os
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image

# Step 1: Merge all Excel files from a folder
folder_path = "./sales_data"  # Put all Excel files in this folder
all_files = glob.glob(os.path.join(folder_path, "*.xlsx"))

combined_df = pd.DataFrame()

for file in all_files:
    df = pd.read_excel(file)
    combined_df = pd.concat([combined_df, df], ignore_index=True)

# Step 2: Clean and summarize data
summary = combined_df.groupby(['Month', 'Product'])['Revenue'].sum().reset_index()
pivot_table = summary.pivot(index='Month', columns='Product', values='Revenue').fillna(0)

# Step 3: Plot the summary chart
plt.figure(figsize=(10,6))
pivot_table.plot(kind='bar', stacked=True)
plt.title('Monthly Revenue by Product')
plt.ylabel('Revenue')
plt.tight_layout()
plt.savefig('report_chart.png')
plt.close()

# Step 4: Create Excel report
wb = Workbook()
ws = wb.active
ws.title = "Sales Report"

# Add summary table
for r in dataframe_to_rows(pivot_table.reset_index(), index=False, header=True):
    ws.append(r)

# Add chart image
img = Image("report_chart.png")
img.anchor = "A20"
ws.add_image(img)

# Step 5: Save final report
wb.save("Auto_Sales_Report.xlsx")

print("✅ Report generated: Auto_Sales_Report.xlsx")
