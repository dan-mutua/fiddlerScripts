import requests
import pandas as pd
import colorama
from colorama import Fore, Style
import json
import socket
from datetime import datetime as dt
import time

colorama.init()

ip_cache = {}

excel_path = r"C:\scripts\filtered_results.xlsx"
df = pd.read_excel(excel_path)

def whois_lookup(ip):
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.connect(("whois.arin.net", 43))
        s.send(('n ' + ip + '\r\n').encode())

        response = b""
        start_time = time.mktime(dt.now().timetuple())
        time_limit = 3
        
        while True:
            elapsed_time = time.mktime(dt.now().timetuple()) - start_time
            data = s.recv(4096)
            response += data
            if (not data) or (elapsed_time >= time_limit):
                break
        
        s.close()
        return response.decode()
    
    except Exception as e:
        print(f"WHOIS lookup error for IP {ip}: {e}")
        return None

def check_ip_ownership(ip):
    try:
        # Check cache first
        if ip in ip_cache:
            print(f"Using cached result for IP {ip}")
            return ip_cache[ip]
            
        # try REST API
        url = f"https://rdap.arin.net/bootstrap/ip/{ip}"
        headers = {'Accept': 'application/json'}
        response = requests.get(url, headers=headers)
        
        result = 'Fail'
        if response.status_code == 200:
            json_data = response.json()
            
            if 'entities' in json_data:
                for entity in json_data['entities']:
                    if 'handle' in entity:
                        name = entity['handle'].upper()
                        if name in ['MSFT', 'AKIMAAI']:
                            result = 'Pass'
                            break
        
        if result == 'Fail':
            # try WHOIS
            whois_response = whois_lookup(ip)
            if whois_response:
                #print(f"\nWHOIS Response for IP {ip}:")
                
                if 'MSFT' in whois_response.upper() or 'AKAMAI' in whois_response.upper():
                    result = 'Pass'
        
        # Cache the result
        ip_cache[ip] = result
        return result
    
    except Exception as e:
        print(f"Error processing IP {ip}: {e}")
        return 'Fail'

df['arinResult'] = df['x-hostip'].apply(check_ip_ownership)

#print
for i, (ip, result) in enumerate(zip(df['x-hostip'].head(20), df['arinResult'].head(20))):
    if result == 'Pass':
        print(f"{ip}: {Fore.GREEN}{result}{Style.RESET_ALL}")
    else:
        print(f"{ip}: {Fore.RED}{result}{Style.RESET_ALL}")

pass_count = (df['arinResult'] == 'Pass').sum()
fail_count = (df['arinResult'] == 'Fail').sum()
print(f"\nSummary: {Fore.GREEN}{pass_count} Pass{Style.RESET_ALL}, {Fore.RED}{fail_count} Fail{Style.RESET_ALL}")

df.to_excel(excel_path, index=False)

print("\nResults have been recorded in the Excel file.")


def combine_results(row):
    if row['arinResult'] == 'Pass' or row['msDocsResult'] == 'Pass':
        return 'Pass'
    return 'Fail'

df['results'] = df.apply(combine_results, axis=1)

final_pass_count = (df['results'] == 'Pass').sum()
final_fail_count = (df['results'] == 'Fail').sum()
print(f"\nFinal Combined Results Summary:")
print(f"{Fore.GREEN}{final_pass_count} Pass{Style.RESET_ALL}, {Fore.RED}{final_fail_count} Fail{Style.RESET_ALL}")

df.to_excel(excel_path, index=False)
print("\nFinal results have been recorded in the Excel file.")