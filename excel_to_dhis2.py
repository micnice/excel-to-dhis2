import requests
import pandas as pd
from requests.auth import HTTPBasicAuth
import json
from datetime import datetime

# Configuration - Update these values
DHIS2_URL = "https://your-dhis2-instance.org/api"
DHIS2_USERNAME = "admin"
DHIS2_PASSWORD = "district"
EXCEL_FILE_PATH = "data.xlsx"
ORG_UNIT_ID = "ID_OF_YOUR_ORGANISATION_UNIT"  # e.g. "abc123"
PERIOD = "202301"  # Format YYYYMM or other DHIS2 period formats
DATA_SET_ID = "ID_OF_YOUR_DATASET"  # Required for aggregate data
ATTR_OPTION_COMBO = "DEFAULT_CATEGORY_COMBO"  # Usually "HllvX50cXC0" for default

def read_excel_data(file_path):
    """Read data from Excel file and return as DataFrame"""
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def prepare_aggregate_payload(df, org_unit, period, data_set, attr_option_combo=None):
    """Convert DataFrame to DHIS2 aggregate API payload format"""
    data_values = []
    
    # If no attribute option combo is provided, use default
    if not attr_option_combo:
        attr_option_combo = "HllvX50cXC0"  # DHIS2 default category option combo
    
    # Iterate through each row in the DataFrame
    for _, row in df.iterrows():
        # Iterate through each column (data element) in the row
        for de_id, value in row.items():
            # Skip null values
            if pd.isna(value):
                continue
                
            # Create data value object for each non-null value
            data_value = {
                "dataElement": de_id,
                "period": period,
                "orgUnit": org_unit,
                "categoryOptionCombo": attr_option_combo,
                "value": str(value),
                "comment": "Imported from Excel"  # Optional comment
            }
            data_values.append(data_value)
    
    payload = {
        "dataSet": data_set,
        "completeDate": datetime.now().strftime("%Y-%m-%d"),
        "period": period,
        "orgUnit": org_unit,
        "dataValues": data_values
    }
    
    return payload

def send_aggregate_to_dhis2(payload, url, username, password):
    """Send aggregate data to DHIS2 instance"""
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    
    auth = HTTPBasicAuth(username, password)
    
    try:
        response = requests.post(
            f"{url}/dataValueSets",
            data=json.dumps(payload),
            headers=headers,
            auth=auth
        )
        
        response.raise_for_status()
        print("Aggregate data successfully sent to DHIS2!")
        print("Response:", response.json())
        return True
        
    except requests.exceptions.RequestException as e:
        print(f"Error sending data to DHIS2: {e}")
        if hasattr(e, 'response') and e.response is not None:
            print("Response content:", e.response.text)
        return False

def main():
    # Read data from Excel
    df = read_excel_data(EXCEL_FILE_PATH)
    if df is None:
        return
    
    # Prepare DHIS2 payload for aggregate data
    payload = prepare_aggregate_payload(df, ORG_UNIT_ID, PERIOD, DATA_SET_ID, ATTR_OPTION_COMBO)
    
    # Send to DHIS2
    send_aggregate_to_dhis2(payload, DHIS2_URL, DHIS2_USERNAME, DHIS2_PASSWORD)

if __name__ == "__main__":
    main()