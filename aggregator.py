import os
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
    Creates a 'Summary_Dashboard' sheet with rich formatting.
    """
    workbook = writer.book
    sheet_name = 'Summary_Dashboard'
    
    # 1. Prepare Data
    # Fail is where Total_Result != 'OK'
    df['is_fail'] = df['Total_Result'].apply(lambda x: 0 if str(x).strip().upper() == 'OK' else 1)
    
    # Placeholders for failure modes (Noise, Index Fail, RPM NG)
    # Checking Intelligent_Control or Section columns for keywords
    df['noise_fail'] = df.apply(lambda row: 1 if 'noise' in str(row.get('Intelligent_Control', '')).lower() or 'noise' in str(row.get('Section', '')).lower() else 0, axis=1)
    df['index_fail'] = df.apply(lambda row: 1 if 'index' in str(row.get('Intelligent_Control', '')).lower() or 'index' in str(row.get('Section', '')).lower() else 0, axis=1)
    df['rpm_fail'] = df.apply(lambda row: 1 if 'rpm' in str(row.get('Intelligent_Control', '')).lower() or 'rpm' in str(row.get('Section', '')).lower() else 0, axis=1)
    df['barcode_fail'] = df.apply(lambda row: 1 if 'barcode' in str(row.get('Intelligent_Control', '')).lower() or 'barcode' in str(row.get('Section', '')).lower() else 0, axis=1)

    lines = sorted(df['Line_Name'].unique())
    
    # Start writing at row 1
    current_row = 1
    
    # Styles
    gold_fill = PatternFill(start_color='FFD700', end_color='FFD700', fill_type='solid')
    blue_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    green_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center')

    # Create sheet if it doesn't exist
    if sheet_name not in workbook.sheetnames:
        workbook.create_sheet(sheet_name)
    ws = workbook[sheet_name]

    for line in lines:
        line_df = df[df['Line_Name'] == line]
        stations = sorted(line_df['Device_ID'].unique())
        
        # Header: Line Name (Merged)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(stations)+2)
        cell = ws.cell(row=current_row, column=1, value=line)
        cell.font = Font(bold=True, size=14)
        cell.alignment = center_align
        current_row += 1

        # Column Headers
        headers = ['Metric', 'Total'] + stations
        for i, h in enumerate(headers):
            c = ws.cell(row=current_row, column=i+1, value=h)
            c.font = header_font
            c.border = thin_border
            c.alignment = center_align
        current_row += 1

        # Metrics definition: (Label, logic_col, fill_style, is_rate)
        metrics = [
            ('Total Count', None, gold_fill, False),
            ('Fail Count', 'is_fail', gold_fill, False),
            ('Fail Rate', 'is_fail', gold_fill, True),
            ('Noise Issues', 'noise_fail', blue_fill, False),
            ('Index Issues', 'index_fail', blue_fill, False),
            ('RPM Issues', 'rpm_fail', green_fill, False),
            ('Barcode Issues', 'barcode_fail', green_fill, False),
        ]

        for label, col, fill, is_rate in metrics:
            # Label cell
            ws.cell(row=current_row, column=1, value=label).border = thin_border
            
            # Line Total
            if label == 'Total Count':
                val = len(line_df)
            elif is_rate:
                total = len(line_df)
                fails = line_df[col].sum()
                val = f"{(fails/total)*100:.2f}%" if total > 0 else "0.00%"
            else:
                val = line_df[col].sum()
            
            c_total = ws.cell(row=current_row, column=2, value=val)
            c_total.fill = fill
            c_total.border = thin_border
            c_total.alignment = center_align

            # Station levels
            for i, station in enumerate(stations):
                st_df = line_df[line_df['Device_ID'] == station]
                if label == 'Total Count':
                    st_val = len(st_df)
                elif is_rate:
                    total = len(st_df)
                    fails = st_df[col].sum()
                    st_val = f"{(fails/total)*100:.2f}%" if total > 0 else "0.00%"
                else:
                    st_val = st_df[col].sum()
                
                c_st = ws.cell(row=current_row, column=i+3, value=st_val)
                c_st.fill = fill
                c_st.border = thin_border
                c_st.alignment = center_align
            
            current_row += 1
        
        current_row += 1 # Spacer between lines

def run_aggregation(target_date=None):
    """
    Aggregate distributed factory log files for a specific date.
    :param target_date: String in 'YYYYMMDD' format. Defaults to yesterday.
    """
    config = configparser.ConfigParser()
    config.read('factory-aggregator/config.ini', encoding='utf-8')

    source_dir = config['Path']['Source_Folder']
    output_dir = config['Path']['Output_Folder'].replace('.\\', 'factory-aggregator/') # Cross-platform tweak
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 1. Handle Date (Default to Yesterday)
    if target_date is None:
        from datetime import timedelta
        target_date = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
    
    logging.info(f"Starting aggregation for date: {target_date}")

    # 2. Map Devices
    device_map = {}
    if 'Device_Mapping' in config:
        for ip, mapping in config['Device_Mapping'].items():
            try:
                line, station = mapping.split(',')
                device_map[ip.replace('.', '_')] = {'Line': line.strip(), 'Station': station.strip()}
            except ValueError:
                logging.error(f"Malformed mapping for IP {ip}: {mapping}")

    # 3. Find relevant files
    search_pattern = os.path.join(source_dir, f"{target_date}_*.txt")
    files = glob.glob(search_pattern)
    
    if not files:
        logging.warning(f"No files found for date {target_date} in {source_dir}")
        return

    all_data = []

    # 4. Process each file
    for file_path in files:
        filename = os.path.basename(file_path)
        match = re.match(rf"{target_date}_(.+)\.txt", filename)
        if not match:
            continue
            
        ip_key = match.group(1)
        meta = device_map.get(ip_key, {'Line': 'Unknown_Line', 'Station': f'Unknown_{ip_key}'})

        try:
            df = pd.read_csv(file_path) 
            df['Line_Name'] = meta['Line']
            df['Device_ID'] = meta['Station']
            df['Source_IP'] = ip_key.replace('_', '.')
            df['Log_Date'] = target_date
            all_data.append(df)
            logging.info(f"Processed: {filename} ({len(df)} rows)")
        except Exception as e:
            logging.error(f"Failed to read {filename}: {e}")

    # 5. Merge and Export to Excel
    if all_data:
        master_df = pd.concat(all_data, ignore_index=True)
        output_file = os.path.join(output_dir, f"Daily_Summary_{target_date}.xlsx")
        
        # Use ExcelWriter with openpyxl engine
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Sheet 1: Raw Data (Named exactly as target_date)
            master_df.to_excel(writer, sheet_name=target_date, index=False)
            
            # Sheet 2: Summary Dashboard
            create_summary_dashboard(writer, master_df, target_date)
            
        logging.info(f"Successfully aggregated {len(files)} files into {output_file}")
        logging.info(f"Total rows: {len(master_df)}")
    else:
        logging.warning("No data was successfully read. Output file not created.")

if __name__ == "__main__":
    run_aggregation()
