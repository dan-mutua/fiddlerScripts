import requests
import pandas as pd
import colorama
from colorama import Fore, Style
import xml.etree.ElementTree as ET

colorama.init()

# Load the Excel file
excel_path = r"C:\scripts\filtered_results.xlsx"
df = pd.read_excel(excel_path)

def check_ip_ownership(ip):
    try:
        # ARIN RDAP Bootstrap API endpoint
        url = f"https://rdap.arin.net/bootstrap/ip/{ip}"
        
        # Headers to request XML response
        headers = {
            'Accept': 'application/xml'
        }
        
        # Send request to ARIN
        response = requests.get(url, headers=headers)
        
        # Check if request was successful
        if response.status_code != 200:
            return 'Fail'
        
        # Parse the XML response
        root = ET.fromstring(response.content)
        
        ns = {'rdap': 'http://rdap.arin.net/registry'}
        
        entity_names = root.findall('.//rdap:entity/rdap:handle', namespaces=ns)
        
        # Check if any entity name matches MSFT or AKIMAAI
        for entity in entity_names:
            name = entity.text.upper()
            if name in ['MSFT', 'AKIMAAI']:
                return 'Pass'
        
        return 'Fail'
    
    except Exception as e:
        print(f"Error processing IP {ip}: {e}")
        return 'Fail'

# Apply the IP ownership check to x-hostip column
df['arinResult'] = df['x-hostip'].apply(check_ip_ownership)

print("\nSample Results (first 20):")
for i, (ip, result) in enumerate(zip(df['x-hostip'].head(20), df['IP_Ownership_Result'].head(20))):
    if result == 'Pass':
        print(f"{ip}: {Fore.GREEN}{result}{Style.RESET_ALL}")
    else:
        print(f"{ip}: {Fore.RED}{result}{Style.RESET_ALL}")

pass_count = (df['arinResult'] == 'Pass').sum()
fail_count = (df['arinResult'] == 'Fail').sum()
print(f"\nSummary: {Fore.GREEN}{pass_count} Pass{Style.RESET_ALL}, {Fore.RED}{fail_count} Fail{Style.RESET_ALL}")

df.to_excel(excel_path, index=False)

print("\nResults have been recorded in the Excel file.")
