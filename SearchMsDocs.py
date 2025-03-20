import requests
import pandas as pd

# Fetch JSON data from the URL
url = "https://endpoints.office.com/endpoints/worldwide?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7"
response = requests.get(url)
data = response.json()

urls = [item['urls'] for item in data if 'urls' in item]
flat_urls = [url for sublist in urls for url in sublist]

excel_path = r"C:\scripts\filtered_results.xlsx"
df = pd.read_excel(excel_path)

df['Result'] = df['https-client-snihostname'].apply(
    lambda hostname: 'Pass' if hostname in flat_urls or 
    ('.' in hostname and hostname[hostname.find('.'):]) in flat_urls else 'Fail'
)

df.to_excel(excel_path, index=False)

print("Results have been recorded in the Excel file.")

