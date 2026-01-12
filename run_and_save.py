import pandas as pd
import os
from processor import process_files

# Setup paths
base_path = '/Users/saritsuriyasangpetch/Documents/KP tools/affiliate-employee'
master_file = os.path.join(base_path, 'example/master/master_Employee_list_08122025.xlsx')
update_dir = os.path.join(base_path, 'example/Dec')
output_dir = os.path.join(base_path, 'output_files')

os.makedirs(output_dir, exist_ok=True)

# Find all update files
update_files = [
    os.path.join(update_dir, f) 
    for f in os.listdir(update_dir) 
    if f.endswith('.xlsx') and not f.startswith('~$')
]

print(f"Processing with Master: {master_file}")
print(f"Found {len(update_files)} update files.")

# Prepare file objects
with open(master_file, 'rb') as f:
    master_content = f.read()

import io
f_master = io.BytesIO(master_content)
f_master.name = 'master_Employee_list_08122025.xlsx'

f_updates = []
for up_path in update_files:
    with open(up_path, 'rb') as f:
        content = f.read()
    obj = io.BytesIO(content)
    obj.name = os.path.basename(up_path)
    f_updates.append(obj)

# Run process
print("Running processor...")
results = process_files(f_master, f_updates)

# Save outputs
print("\nSaving files to 'output_files/' directory:")
for name, content in results.items():
    out_path = os.path.join(output_dir, name)
    with open(out_path, 'wb') as f:
        f.write(content.read())
    print(f"- Saved {out_path}")

print("\nDone!")
