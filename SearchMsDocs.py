import requests
import pandas as pd
import colorama
from colorama import Fore, Style

colorama.init()

url = "https://endpoints.office.com/endpoints/worldwide?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7"
response = requests.get(url)
data = response.json()

urls = [item['urls'] for item in data if 'urls' in item]
flat_urls = [url.lower() for sublist in urls for url in sublist]

#print(f"Sample URLs from API (first 5): {flat_urls[:5]}")
print(f"Total URLs fetched: {len(flat_urls)}")

excel_path = r"C:\scripts\filtered_results.xlsx"
df = pd.read_excel(excel_path)

def check_hostname(hostname):
    if not isinstance(hostname, str):
        return 'Fail'
    
    hostname = hostname.lower()
    
    if any(hostname in url for url in flat_urls):
        return 'Pass'
    
    dot_positions = []
    start_index = 0
    
    for _ in range(3):
        dot_index = hostname.find('.', start_index)
        if dot_index == -1:
            break
        dot_positions.append(dot_index)
        start_index = dot_index + 1
    
    for dot_pos in dot_positions:
        domain_part = hostname[dot_pos:]
        if any(domain_part in url for url in flat_urls):
            return 'Pass'
    
    return 'Fail'

df['msDocsResult'] = df['https-client-snihostname'].apply(check_hostname)



pass_count = (df['msDocsResult'] == 'Pass').sum()
fail_count = (df['msDocsResult'] == 'Fail').sum()
print(f"\nSummary: {Fore.GREEN}{pass_count} Pass{Style.RESET_ALL}, {Fore.RED}{fail_count} Fail{Style.RESET_ALL}")

df.to_excel(excel_path, index=False)

print("\nResults have been recorded in the Excel file.")