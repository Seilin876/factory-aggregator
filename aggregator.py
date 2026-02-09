import os
import pandas as pd
import configparser
import logging
from datetime import datetime
import glob
import re

# Initialize logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def run_aggregation(target_date=None):
    """
    Aggregate distributed factory log files for a specific date.
    :param target_date: String in 'YYYYMMDD' format. Defaults to yesterday.
    """
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')

    source_dir = config['Path']['Source_Folder']
    output_dir = config['Path']['Output_Folder']
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 1. Handle Date (Default to Yesterday)
    if target_date is None:
        from datetime import timedelta
        target_date = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
    
    logging.info(f"Starting aggregation for date: {target_date}")

    # 2. Map Devices
    device_map = {}
    for ip, mapping in config['Device_Mapping'].items():
        try:
            line, station = mapping.split(',')
            device_map[ip.replace('.', '_')] = {'Line': line.strip(), 'Station': station.strip()}
        except ValueError:
            logging.error(f"Malformed mapping for IP {ip}: {mapping}")

    # 3. Find relevant files
    # Pattern: {Date}_{IP}.txt
    search_pattern = os.path.join(source_dir, f"{target_date}_*.txt")
    files = glob.glob(search_pattern)
    
    if not files:
        logging.warning(f"No files found for date {target_date} in {source_dir}")
        return

    all_data = []

    # 4. Process each file
    for file_path in files:
        filename = os.path.basename(file_path)
        # Extract IP from filename (Target: {Date}_{IP}.txt)
        match = re.match(rf"{target_date}_(.+)\.txt", filename)
        if not match:
            continue
            
        ip_key = match.group(1)
        meta = device_map.get(ip_key, {'Line': 'Unknown_Line', 'Station': f'Unknown_{ip_key}'})

        try:
            # Assuming CSV/Text format. Adjust separator if necessary (defaulting to comma)
            df = pd.read_csv(file_path) 
            
            # Inject Metadata
            df['Line_Name'] = meta['Line']
            df['Device_ID'] = meta['Station']
            df['Source_IP'] = ip_key.replace('_', '.')
            df['Log_Date'] = target_date
            
            all_data.append(df)
            logging.info(f"Processed: {filename} ({len(df)} rows)")
        except Exception as e:
            logging.error(f"Failed to read {filename}: {e}")

    # 5. Merge and Export
    if all_data:
        master_df = pd.concat(all_data, ignore_index=True)
        output_file = os.path.join(output_dir, f"Daily_Summary_{target_date}.csv")
        master_df.to_csv(output_file, index=False)
        logging.info(f"Successfully aggregated {len(files)} files into {output_file}")
        logging.info(f"Total rows: {len(master_df)}")
    else:
        logging.warning("No data was successfully read. Output file not created.")

if __name__ == "__main__":
    # Example: To run for a specific date, pass '20260208'
    run_aggregation()
