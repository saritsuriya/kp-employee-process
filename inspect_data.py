import pandas as pd
import os

base_path = '/Users/saritsuriyasangpetch/Documents/KP tools/affiliate-employee'
update_file = os.path.join(base_path, 'example/Dec/01.KPG-ข้อมูลพนักงาน สำหรับกิจกรรม AP_Dec 2025 Data as of 01.12.2025.xlsx')

pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

print(f"\n--- Dumping Rows for Update File: {update_file} ---")
# Read first 20 rows without header
df = pd.read_excel(update_file, sheet_name='Current', header=None, nrows=20)
print(df)

print("\n--- Dumping Rows for Resign Tab ---")
df_resign = pd.read_excel(update_file, sheet_name='Resign', header=None, nrows=20)
print(df_resign)
