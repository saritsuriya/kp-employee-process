import pandas as pd
import io

def normalize_columns(df):
    """Normalize column names to strip whitespace and handle potential variations."""
    df.columns = df.columns.astype(str).str.strip()
    return df

def process_files(master_file, update_files):
    """
    Process the master file and a list of update files.
    
    Args:
        master_file: Uploaded file object for the Master Employee List.
        update_files: List of Uploaded file objects for the HR Update Lists.
        
    Returns:
        Dictionary containing BytesIO objects for the generated files:
        - 'sap_file.xlsx'
        - 'resigned_batch_file.xlsx'
        - 'new_master_file.xlsx'
    """
    
    # --- 1. Load Master File ---
    # Master file has header at row 0 (default)
    df_master = pd.read_excel(master_file, header=0)
    df_master = normalize_columns(df_master)
    
    # Ensure key columns exist (fallback to index if names slightly vary, but assuming inspection was correct)
    # Master Columns: 'Employee No.', 'Employee Name', etc.
    # Convert Employee No. to string to ensure matching works
    df_master['Employee No.'] = df_master['Employee No.'].apply(lambda x: str(int(x)) if pd.notnull(x) and isinstance(x, (int, float)) else str(x).strip())
    
    # --- 2. Process Update Files ---
    new_hires_list = []
    resigned_list = []
    
    for update_file in update_files:
        # Load Current Tab (Index 3 header)
        try:
            df_current = pd.read_excel(update_file, sheet_name='Current', header=3)
            df_current = normalize_columns(df_current)
            # Find Employee No. column in Current tab
            # Current Tab Columns: 'Employee No.', 'Employee Name', etc.
            if 'Employee No.' in df_current.columns:
                 # Drop rows where Employee No is NaN or "No data" (case insensitive)
                 df_current = df_current[df_current['Employee No.'].notna()]
                 df_current = df_current[~df_current['Employee No.'].astype(str).str.contains(r'No data', case=False, na=False)]
                 
                 # Drop rows where Employee No might be empty string after strip
                 df_current['Employee No.'] = df_current['Employee No.'].apply(lambda x: str(int(x)) if pd.notnull(x) and isinstance(x, (int, float)) else str(x).strip())
                 df_current = df_current[df_current['Employee No.'] != '']
                 
                 # Identify New Hires: In Update but NOT in Master
                 # We filter df_current where Employee No. is NOT in df_master['Employee No.']
                 new_hires = df_current[~df_current['Employee No.'].isin(df_master['Employee No.'])]
                 new_hires_list.append(new_hires)
            
        except Exception as e:
            print(f"Error processing 'Current' tab in {update_file.name}: {e}")

        # Load Resign Tab (Index 3 header)
        try:
            df_resign = pd.read_excel(update_file, sheet_name='Resign', header=3)
            df_resign = normalize_columns(df_resign)
            
            # Filter out empty rows or instruction rows if any slip through (though header=3 should catch most)
            if 'Employee No.' in df_resign.columns:
                 # Drop rows where Employee No is NaN or "No data" (case insensitive)
                 df_resign = df_resign[df_resign['Employee No.'].notna()]
                 df_resign = df_resign[~df_resign['Employee No.'].astype(str).str.contains(r'No data', case=False, na=False)]
                 
                 # Drop rows where Employee No might be empty string after strip
                 df_resign['Employee No.'] = df_resign['Employee No.'].apply(lambda x: str(int(x)) if pd.notnull(x) and isinstance(x, (int, float)) else str(x).strip())
                 df_resign = df_resign[df_resign['Employee No.'] != '']
                 
                 resigned_list.append(df_resign)
                 
        except Exception as e:
            print(f"Error processing 'Resign' tab in {update_file.name}: {e}")

    # --- 3. Consolidate Lists ---
    if new_hires_list:
        df_all_new = pd.concat(new_hires_list, ignore_index=True)
    else:
        df_all_new = pd.DataFrame() # Empty schema if needed, or just empty
        
    if resigned_list:
        df_all_resigned = pd.concat(resigned_list, ignore_index=True)
    else:
        df_all_resigned = pd.DataFrame()

    # --- 4. Generate Output Dataframes ---

    # Output A: SAP File (Full Master + New Hires + Resignation Dates)
    # 1. Start with Master
    df_sap = df_master.copy()
    
    # 2. Add 'Resignation Date' column if not exists
    if 'Resignation Date' not in df_sap.columns:
        df_sap['Resignation Date'] = None
    
    # 3. Update Resignation Dates for Resigned Employees
    if not df_all_resigned.empty:
        # Create a mapping dictionary: Employee No. -> Resignation Date
        # Ensure dates are datetime objects or strings as needed
        resign_map = dict(zip(df_all_resigned['Employee No.'], df_all_resigned['วันที่พ้นสภาพ']))
        
        # Apply mapping
        # We iterate or use map. Using map on a subset is efficient.
        # Find index of employees who resigned
        resigned_mask = df_sap['Employee No.'].isin(resign_map.keys())
        
        # Update the 'Resignation Date' for these rows
        # We use apply to look up the date in the map
        df_sap.loc[resigned_mask, 'Resignation Date'] = df_sap.loc[resigned_mask, 'Employee No.'].map(resign_map)

    # 4. Append New Hires
    if not df_all_new.empty:
        # We need to map New Hire columns to Master columns, similar to New Master logic
        # Create a temporary dataframe for new hires with master columns
        df_new_hires_mapped = pd.DataFrame(columns=df_sap.columns)
        
        # Map known columns
        df_new_hires_mapped['Employee No.'] = df_all_new['Employee No.']
        
        # Map matching columns
        common_cols = list(set(df_sap.columns) & set(df_all_new.columns))
        for col in common_cols:
             df_new_hires_mapped[col] = df_all_new[col]
             
        # Append to SAP dataframe
        df_sap = pd.concat([df_sap, df_new_hires_mapped], ignore_index=True)
    
    # Final cleanup: Re-index 'No.' column
    if 'No.' in df_sap.columns:
        df_sap['No.'] = range(1, len(df_sap) + 1)
    
    # Output B: Batch Resigned File
    df_batch_resigned = df_all_resigned.copy()
    
    # Output C: New Master List
    # Start with original master
    df_new_master = df_master.copy()
    
    # Remove Resigned
    if not df_batch_resigned.empty:
        resigned_ids = df_batch_resigned['Employee No.'].unique()
        df_new_master = df_new_master[~df_new_master['Employee No.'].isin(resigned_ids)]
        
    # Add New Hires
    # We need to map New Hire columns (from Update file) to Master columns
    # Master Cols: No., Employee No., Employee Name, Employee Name Eng, Hire date, Company, Group, Division, Department, Location, AP Code
    # Update Cols (approx): No., Employee No., Employee Name Thai, Company, ...
    
    # Mapping attempt based on column names from inspection
    if not df_all_new.empty:
        # Create a dataframe with Master columns
        df_to_add = pd.DataFrame(columns=df_master.columns)
        
        # Map known columns
        # Update file: 'Employee No.' -> Master: 'Employee No.'
        df_to_add['Employee No.'] = df_all_new['Employee No.']
        
        # Update: 'Employee Name Thai' or just 'Employee Name' (need to check what inspection showed)
        # Inspection showed: 'Employee Name Thai' in Resign, but 'Employee Name' in Current?
        # Let's check logic: Current tab inspection showed columns: 'No.', 'Employee No.', 'Employee Name', 'Employee Name Eng', 'Hire date'... matches Master!
        # So we can likely just do a direct concat if names match.
        
        common_cols = list(set(df_master.columns) & set(df_all_new.columns))
        for col in common_cols:
            df_to_add[col] = df_all_new[col]
            
        # Append
        df_new_master = pd.concat([df_new_master, df_to_add], ignore_index=True)
        
    # Re-index 'No.' column
    if 'No.' in df_new_master.columns:
        df_new_master['No.'] = range(1, len(df_new_master) + 1)

    # Helper to clean date columns
    def clean_date_columns(df, cols):
        for col in cols:
            if col in df.columns:
                # Convert to datetime first to handle mixed formats, then format to string YYYY-MM-DD
                # If NaT, keep as None/NaN
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')
    
    # Helper to generate AP Code
    def generate_ap_code(df):
        if 'Employee No.' in df.columns:
            # Ensure Employee No. is cleaned string
            clean_ids = df['Employee No.'].astype(str).str.strip()
            df['AP Code'] = 'AP' + clean_ids
        return df

    # --- Apply Logic to Outputs ---

    # 1. SAP File
    df_sap = generate_ap_code(df_sap)
    clean_date_columns(df_sap, ['Hire date', 'Resignation Date', 'วันที่พ้นสภาพ']) # Add other date columns if needed

    # 2. Batch Resigned File
    df_batch_resigned = generate_ap_code(df_batch_resigned)
    clean_date_columns(df_batch_resigned, ['Hire date', 'Resignation Date', 'วันที่พ้นสภาพ'])
    
    # Re-index 'No.' column for batch file to be continuous
    if 'No.' in df_batch_resigned.columns:
        df_batch_resigned['No.'] = range(1, len(df_batch_resigned) + 1)

    # 3. New Master File
    df_new_master = generate_ap_code(df_new_master)
    clean_date_columns(df_new_master, ['Hire date', 'Resignation Date', 'วันที่พ้นสภาพ'])

    # --- 5. Return BytesIO objects ---
    output_files = {}

    def to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        return output

    output_files['sap_file.xlsx'] = to_excel(df_sap)
    output_files['resigned_batch_file.xlsx'] = to_excel(df_batch_resigned)
    output_files['new_master_file.xlsx'] = to_excel(df_new_master)

    return output_files
