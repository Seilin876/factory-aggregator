import os
import sys
import pandas as pd
import configparser
import logging
from datetime import datetime
import glob
import re
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

# Initialize logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def create_summary_dashboard(writer, df, date_str):
    """
    Creates a 'Summary_Dashboard' sheet with rich formatting based on Delta Electronics brand colors
    and specific Grouping/Ranking logic.
    """
    workbook = writer.book
    sheet_name = 'Summary_Dashboard'
    
    # --- 1. Data Pre-processing & Logic Implementation ---
    
    # A. Convert critical columns to Numeric
    numeric_cols = ['index1', 'index1_limit', 'index2', 'index2_limit', 'RPM', 'RPM_Low']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        else:
            df[col] = 0

    # B. Define Basic Conditions
    cond_rotating = (df['RPM'] != 0)
    cond_idx1_high = (df['index1'] > df['Index1_Limit'])
    cond_idx2_high = (df['index2'] > df['Index2_Limit'])
    cond_out_of_control = (df['RPM'] > df['RPM_Low']) + (df['RPM'] > 10000)
    cond_no_rotate = (df['RPM'] == 0)

    # --- 2. Calculate Final Categories (0 or 1) ---

    # 1. Total Fail
    df['is_fail'] = df['Total_Result'].apply(lambda x: 0 if str(x).strip().upper() == 'OK' else 1)

    # 2. Noise Logic
    df['calc_noise'] = (
        ((cond_idx1_high) & (~cond_idx2_high) & (cond_rotating)).astype(int) +
        ((cond_idx2_high) & (~cond_idx1_high) & (cond_rotating)).astype(int) +
        ((cond_idx1_high) & (cond_idx2_high) & (cond_rotating)).astype(int) +
        df.apply(lambda row: 1 if str(row.get('Intelligent_Control','')).strip().upper() == 'OK' 
                 and str(row.get('Total_Result','')).strip().upper() != 'OK' else 0, axis=1)
    )

    # 3. Only Index1 Fail
    df['calc_only_idx1'] = ((cond_idx1_high) & (~cond_idx2_high) & (cond_rotating)).astype(int)

    # 4. Only Index2 Fail
    df['calc_only_idx2'] = ((cond_idx2_high) & (~cond_idx1_high) & (cond_rotating)).astype(int)

    # 5. Index 1 & 2 both Fail
    df['calc_both_idx'] = ((cond_idx1_high) & (cond_idx2_high) & (cond_rotating)).astype(int)

    # 6. Spec Fail
    df['calc_spec_fail'] = df.apply(lambda row: 1 if str(row.get('Intelligent_Control','')).strip().upper() == 'OK' 
                                    and str(row.get('Total_Result','')).strip().upper() != 'OK' else 0, axis=1)

    # 7. Out of control
    df['calc_out_control'] = cond_out_of_control.astype(int)

    # 8. No rotate
    df['calc_no_rotate'] = cond_no_rotate.astype(int)

    # 9. RPM NG
    df['calc_rpm_ng'] = ((cond_out_of_control) | (cond_no_rotate)).astype(int)

    # 10. PauseOrFreeRun
    df['calc_pause'] = df['Model_Name'].apply(lambda x: 1 if 'pauseorfreerun' in str(x).lower().replace(" ", "") else 0)

    # 11. No Barcode
    df['calc_no_barcode'] = df['Barcode'].apply(lambda x: 1 if pd.isna(x) or str(x).strip() == '' else 0)

    # 12. Others
    df['calc_others'] = ((df['calc_pause'] == 1) | (df['calc_no_barcode'] == 1)).astype(int)


    # --- 3. Dashboard Generation ---
    lines = sorted(df['Line_Name'].unique())
    current_row = 1
    
    # --- Define Styles (Delta Electronics Brand Colors & Ranking Colors) ---
    
    # Group 1: Summary (Delta Blue)
    # Main: Deep Blue, Text: White
    fill_g1_main = PatternFill(start_color='0066A1', end_color='0066A1', fill_type='solid') 
    font_g1_main = Font(bold=True, color='FFFFFF')
    # Sub: Very Light Blue
    fill_g1_sub = PatternFill(start_color='E6F2FF', end_color='E6F2FF', fill_type='solid') 
    font_sub_default = Font(bold=False, color='000000')

    # Group 2: Noise (Cyan/Sky)
    # Main: Medium Cyan
    fill_g2_main = PatternFill(start_color='4BB4E6', end_color='4BB4E6', fill_type='solid')
    font_g2_main = Font(bold=True, color='000000')
    # Sub: Light Cyan
    fill_g2_sub = PatternFill(start_color='D9F2FF', end_color='D9F2FF', fill_type='solid')

    # Group 3: RPM (Delta Green)
    # Main: Delta Green
    fill_g3_main = PatternFill(start_color='8CC63F', end_color='8CC63F', fill_type='solid')
    font_g3_main = Font(bold=True, color='000000')
    # Sub: Light Green
    fill_g3_sub = PatternFill(start_color='F0F9E8', end_color='F0F9E8', fill_type='solid')

    # Group 4: Others (Neutral/Grey)
    # Main: Medium Grey
    fill_g4_main = PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid')
    font_g4_main = Font(bold=True, color='000000')
    # Sub: Light Grey
    fill_g4_sub = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')

    # Ranking Colors (Comfortable Pastels)
    rank_1_fill = PatternFill(start_color='FF9999', end_color='FF9999', fill_type='solid') # Light Red
    rank_2_fill = PatternFill(start_color='FFCC99', end_color='FFCC99', fill_type='solid') # Light Orange
    rank_3_fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid') # Light Yellow

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center')

    if sheet_name not in workbook.sheetnames:
        workbook.create_sheet(sheet_name)
    ws = workbook[sheet_name]

    for line in lines:
        line_df = df[df['Line_Name'] == line]
        stations = sorted(line_df['Device_ID'].unique())
        
        # Header: Line Name
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(stations)+2)
        cell = ws.cell(row=current_row, column=1, value=line)
        cell.font = Font(bold=True, size=14, color='FFFFFF')
        cell.fill = fill_g1_main # Use Group 1 Main color for Line Header
        cell.alignment = center_align
        current_row += 1

        # Column Headers
        headers = ['Metric', 'Total'] + stations
        for i, h in enumerate(headers):
            c = ws.cell(row=current_row, column=i+1, value=h)
            c.font = Font(bold=True)
            c.fill = PatternFill(start_color='DDDDDD', end_color='DDDDDD', fill_type='solid') # Grey Header
            c.border = thin_border
            c.alignment = center_align
        current_row += 1

        # --- Metrics Configuration ---
        # Format: (Label, Column_Name, Is_Rate, Is_String, Group_ID, Is_Main_Row)
        # Group IDs: 1=Summary, 2=Noise, 3=RPM, 4=Others
        metrics_config = [
            # Group 1: Summary
            ('Total Count',       None,               False, False, 1, False), # Sub
            ('Fail Count',        'is_fail',          False, False, 1, False), # Sub (Ranked)
            ('Fail Rate',         'is_fail',          True,  False, 1, False), # Sub (Ranked)
            ('Noise Rate',        'calc_noise',       True,  False, 1, False), # Sub (New, Ranked)
            ('RPM Fail Rate',     'calc_rpm_ng',      True,  False, 1, False), # Sub (New, Ranked)
            ('Other Fail Rate',   'calc_others',      True,  False, 1, False), # Sub (New, Ranked)
            
            # Group 2: Noise Fail
            ('Noise',             'calc_noise',       False, False, 2, True),  # Main
            ('Only Index1 Fail',  'calc_only_idx1',   False, False, 2, False),
            ('Only Index2 Fail',  'calc_only_idx2',   False, False, 2, False),
            ('Index 1 & 2 both',  'calc_both_idx',    False, False, 2, False),
            ('Spec Fail',         'calc_spec_fail',   False, False, 2, False),

            # Group 3: RPM Fail
            ('RPM NG',            'calc_rpm_ng',      False, False, 3, True),  # Main
            ('Out of control',    'calc_out_control', False, False, 3, False),
            ('No rotate',         'calc_no_rotate',   False, False, 3, False),

            # Group 4: Others Fail
            ('Others',            'calc_others',      False, False, 4, True),  # Main
            ('PauseOrFreeRun',    'calc_pause',       False, False, 4, False),
            ('No Barcode',        'calc_no_barcode',  False, False, 4, False),
            ('Model Name',        'Model_Name',       False, True,  4, False), # Special case
        ]

        # Rows that require Ranking (Red/Orange/Yellow)
        ranking_targets = ['Fail Count', 'Fail Rate', 'Noise Rate', 'RPM Fail Rate', 'Other Fail Rate']

        for label, col, is_rate, is_string, group_id, is_main in metrics_config:
            
            # Determine Base Style based on Group
            if group_id == 1:
                base_fill = fill_g1_main if is_main else fill_g1_sub
                base_font = font_g1_main if is_main else font_sub_default
            elif group_id == 2:
                base_fill = fill_g2_main if is_main else fill_g2_sub
                base_font = font_g2_main if is_main else font_sub_default
            elif group_id == 3:
                base_fill = fill_g3_main if is_main else fill_g3_sub
                base_font = font_g3_main if is_main else font_sub_default
            else: # Group 4
                base_fill = fill_g4_main if is_main else fill_g4_sub
                base_font = font_g4_main if is_main else font_sub_default

            # Metric Label Cell
            c_label = ws.cell(row=current_row, column=1, value=label)
            c_label.fill = base_fill
            c_label.font = base_font
            c_label.border = thin_border

            # Calculation & Data Storage for Ranking
            station_values = [] # To store (col_index, value) for ranking
            
            # --- Total Column ---
            if is_string:
                val = ",".join(line_df[col].dropna().unique().astype(str))
            elif label == 'Total Count':
                val = len(line_df)
            elif is_rate:
                total = len(line_df)
                fails = line_df[col].sum()
                val = f"{(fails/total)*100:.2f}%" if total > 0 else "0.00%"
            else:
                val = line_df[col].sum()
            
            c_total = ws.cell(row=current_row, column=2, value=val)
            c_total.fill = base_fill
            c_total.font = base_font
            c_total.border = thin_border
            c_total.alignment = center_align

            # --- Station Columns ---
            for i, station in enumerate(stations):
                st_df = line_df[line_df['Device_ID'] == station]
                
                # Calculate Value
                raw_num_val = 0 # For ranking comparison
                
                if is_string:
                    unique_names = st_df[col].dropna().unique()
                    st_val = str(unique_names[0]) if len(unique_names) > 0 else ""
                elif label == 'Total Count':
                    raw_num_val = len(st_df)
                    st_val = raw_num_val
                elif is_rate:
                    total = len(st_df)
                    fails = st_df[col].sum()
                    raw_num_val = (fails/total) if total > 0 else 0
                    st_val = f"{raw_num_val*100:.2f}%"
                else:
                    raw_num_val = st_df[col].sum()
                    st_val = raw_num_val
                
                # Write Cell
                c_st = ws.cell(row=current_row, column=i+3, value=st_val)
                c_st.fill = base_fill # Default to group color
                c_st.font = base_font
                c_st.border = thin_border
                c_st.alignment = center_align
                
                # Store for ranking if needed (col_idx, value)
                # Only rank if value > 0 to avoid coloring empty 0s if desired, 
                # but usually ranking includes 0 if others are 0. Let's strictly rank numbers.
                station_values.append({'col_idx': i+3, 'val': raw_num_val, 'cell': c_st})

            # --- Apply Ranking Logic ---
            if label in ranking_targets:
                # Extract values to find top 3 unique values
                # We only care about values > 0 usually, but if all are valid, rank them.
                # Let's simple rank all numerical values.
                
                # Get unique values sorted descending
                unique_vals = sorted(list(set([x['val'] for x in station_values])), reverse=True)
                
                # Define Rank thresholds
                val_1st = unique_vals[0] if len(unique_vals) > 0 else None
                val_2nd = unique_vals[1] if len(unique_vals) > 1 else None
                val_3rd = unique_vals[2] if len(unique_vals) > 2 else None

                for item in station_values:
                    # Skip ranking if value is 0 (optional, but keeps chart clean)
                    if item['val'] == 0:
                        continue
                        
                    if item['val'] == val_1st:
                        item['cell'].fill = rank_1_fill # Red
                    elif item['val'] == val_2nd:
                        item['cell'].fill = rank_2_fill # Orange
                    elif item['val'] == val_3rd:
                        item['cell'].fill = rank_3_fill # Yellow

            current_row += 1
        
        current_row += 1 # Spacer

def run_aggregation(target_date=None):
    """
    Aggregate distributed factory log files.
    """
    # 1. Robust Config Path Resolution
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    config_path = os.path.join(base_dir, 'config.ini')

    # Config Check
    if not os.path.exists(config_path):
        logging.error(f"Config file NOT found at: {config_path}")
        return

    config = configparser.ConfigParser()
    config.read(config_path, encoding='utf-8')

    try:
        source_dir = config['Path']['Source_Folder']
        output_base = config['Path']['Output_Folder']
        # Handle relative output paths
        if output_base.startswith('.\\') or output_base.startswith('./'):
            output_dir = os.path.join(base_dir, output_base.replace('.\\', '').replace('./', ''))
        else:
            output_dir = output_base
    except KeyError as e:
        logging.error(f"Config file missing key: {e}")
        return
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 2. Handle Date
    if target_date is None:
        from datetime import timedelta
        target_date = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
    
    logging.info(f"Starting aggregation for date: {target_date}")

    # 3. Map Devices
    device_map = {}
    if 'Device_Mapping' in config:
        for ip, mapping in config['Device_Mapping'].items():
            try:
                line, station = mapping.split(',')
                device_map[ip.replace('.', '_')] = {'Line': line.strip(), 'Station': station.strip()}
            except ValueError:
                logging.error(f"Malformed mapping for IP {ip}: {mapping}")

    # 4. Search Files
    search_pattern = os.path.join(source_dir, f"{target_date}_*.txt")
    files = glob.glob(search_pattern)
    
    if not files:
        logging.warning(f"No files found for date {target_date} in {source_dir}")
        return

    all_data = []

    # 5. Process Files
    for file_path in files:
        filename = os.path.basename(file_path)
        match = re.match(rf"{target_date}_(.+)\.txt", filename)
        if not match:
            continue
            
        ip_key = match.group(1)
        meta = device_map.get(ip_key, {'Line': 'Unknown_Line', 'Station': f'Unknown_{ip_key}'})

        try:
            # Robust Reading: Tab separator, CP950 encoding, Skip bad lines, No index column
            df = pd.read_csv(file_path, sep='\t', encoding='cp950', on_bad_lines='skip', index_col=False)
            
            # Header Cleaning (Trim whitespace)
            df.columns = df.columns.str.strip()
            
            # Check essential column
            if 'Total_Result' not in df.columns:
                logging.warning(f"Skipping {filename}: Missing 'Total_Result'")
                continue

            df['Line_Name'] = meta['Line']
            df['Device_ID'] = meta['Station']
            df['Source_IP'] = ip_key.replace('_', '.')
            df['Log_Date'] = target_date
            
            # Convert Result to string to be safe
            df['Total_Result'] = df['Total_Result'].astype(str)
            
            all_data.append(df)
            logging.info(f"Processed: {filename} ({len(df)} rows)")

        except UnicodeDecodeError:
            # Fallback to UTF-8
            try:
                df = pd.read_csv(file_path, sep='\t', encoding='utf-8', on_bad_lines='skip', index_col=False)
                df.columns = df.columns.str.strip()
                df['Line_Name'] = meta['Line']
                df['Device_ID'] = meta['Station']
                df['Source_IP'] = ip_key.replace('_', '.')
                df['Log_Date'] = target_date
                df['Total_Result'] = df['Total_Result'].astype(str)
                all_data.append(df)
                logging.info(f"Processed (UTF-8): {filename}")
            except Exception as e2:
                logging.error(f"Failed {filename} with UTF-8: {e2}")

        except Exception as e:
            logging.error(f"Failed to read {filename}: {e}")

    # 6. Export
    if all_data:
        master_df = pd.concat(all_data, ignore_index=True)
        output_file = os.path.join(output_dir, f"Daily_Summary_{target_date}.xlsx")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Sheet 1: Raw Data
            master_df.to_excel(writer, sheet_name=target_date, index=False)
            
            # Sheet 2: Dashboard
            create_summary_dashboard(writer, master_df, target_date)
            
        logging.info(f"Aggregation Complete. File: {output_file}")
    else:
        logging.warning("No valid data found.")

if __name__ == "__main__":
    run_aggregation()
