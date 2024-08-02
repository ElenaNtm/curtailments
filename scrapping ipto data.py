# -*- coding: utf-8 -*-
"""
Created on Tue Jul 16 16:02:54 2024

@author: Eleni
"""


"""
Download additional data
"""
import requests
import json
import os

"""
Load forecast
"""
url = "https://www.admie.gr/getOperationMarketFilewRange?dateStart=2024-01-01&dateEnd=2024-07-25&FileCategory=isp3intradayloadforecast"

download_dir = r"C:\users\eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\1. load forecast ipto\Intraday"
if not os.path.exists(download_dir):
    os.mkdir(download_dir)

# Send an HTTP GET request to the URL
response = requests.get(url)

if response.status_code == 200:
    # Parse the JSON response as a list
    files = json.loads(response.text)

    for file_info in files:
        file_url = file_info.get("file_path")
        filename = file_info.get("file_description")

        if filename is not None:
            filename = os.path.join(download_dir, filename + ".xlsx")

            # Download the file
            file_response = requests.get(file_url)
            if file_response.status_code == 200:
                with open(filename, 'wb') as file:
                    file.write(file_response.content)
                print(f"Downloaded: {filename}")

            else:
                print(f"Failed to download: {filename}")
        else:
            print("Skipped a file with a None name.")
else:
    print(f"Failed to retrieve JSON data. Status code: {response.status_code}")    



url = "https://www.admie.gr/getOperationMarketFilewRange?dateStart=2024-01-01&dateEnd=2024-07-25&FileCategory=isp2dayaheadloadforecast"

download_dir = r"C:\users\eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\1. load forecast ipto\DA"
if not os.path.exists(download_dir):
    os.mkdir(download_dir)

# Send an HTTP GET request to the URL
response = requests.get(url)

if response.status_code == 200:
    # Parse the JSON response as a list
    files = json.loads(response.text)

    for file_info in files:
        file_url = file_info.get("file_path")
        filename = file_info.get("file_description")

        if filename is not None:
            filename = os.path.join(download_dir, filename + ".xlsx")

            # Download the file
            file_response = requests.get(file_url)
            if file_response.status_code == 200:
                with open(filename, 'wb') as file:
                    file.write(file_response.content)
                print(f"Downloaded: {filename}")

            else:
                print(f"Failed to download: {filename}")
        else:
            print("Skipped a file with a None name.")
else:
    print(f"Failed to retrieve JSON data. Status code: {response.status_code}")       
    
"""
Production-consumption
"""
# URL of the webpage that provides the JSON data
url = "https://www.admie.gr/getOperationMarketFilewRange?dateStart=2024-07-15&dateEnd=2024-07-25&FileCategory=SystemRealizationSCADA"

download_dir = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\4. production-consumption ipto"

if not os.path.exists(download_dir):
    os.mkdir(download_dir)

# Send an HTTP GET request to the URL
response = requests.get(url)

if response.status_code == 200:
    # Parse the JSON response as a list
    files = json.loads(response.text)

    for file_info in files:
        file_url = file_info.get("file_path")
        filename = file_info.get("file_description")

        if filename is not None:
            filename = os.path.join(download_dir, filename + ".xls")

            # Download the file
            file_response = requests.get(file_url)
            if file_response.status_code == 200:
                with open(filename, 'wb') as file:
                    file.write(file_response.content)
                print(f"Downloaded: {filename}")

            else:
                print(f"Failed to download: {filename}")
        else:
            print("Skipped a file with a None name.")
else:
    print(f"Failed to retrieve JSON data. Status code: {response.status_code}")


"""
DA RES production forecast
"""
url = "https://www.admie.gr/getOperationMarketFilewRange?dateStart=2024-01-01&dateEnd=2024-07-25&FileCategory=isp2dayaheadresforecast"

download_dir = r"C:\users\eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\2. DA RES forecasts ipto"
if not os.path.exists(download_dir):
    os.mkdir(download_dir)

# Send an HTTP GET request to the URL
response = requests.get(url)

if response.status_code == 200:
    # Parse the JSON response as a list
    files = json.loads(response.text)

    for file_info in files:
        file_url = file_info.get("file_path")
        filename = file_info.get("file_description")

        if filename is not None:
            filename = os.path.join(download_dir, filename + ".xlsx")

            # Download the file
            file_response = requests.get(file_url)
            if file_response.status_code == 200:
                with open(filename, 'wb') as file:
                    file.write(file_response.content)
                print(f"Downloaded: {filename}")

            else:
                print(f"Failed to download: {filename}")
        else:
            print("Skipped a file with a None name.")
else:
    print(f"Failed to retrieve JSON data. Status code: {response.status_code}")     
    
"""
Intraday RES production forecast
"""
url = "https://www.admie.gr/getOperationMarketFilewRange?dateStart=2024-01-01&dateEnd=2024-07-25&FileCategory=isp3intradayresforecast"

download_dir = r"C:\users\eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\3. Intraday RES forecasts ipto"
if not os.path.exists(download_dir):
    os.mkdir(download_dir)

# Send an HTTP GET request to the URL
response = requests.get(url)

if response.status_code == 200:
    # Parse the JSON response as a list
    files = json.loads(response.text)

    for file_info in files:
        file_url = file_info.get("file_path")
        filename = file_info.get("file_description")

        if filename is not None:
            filename = os.path.join(download_dir, filename + ".xlsx")

            # Download the file
            file_response = requests.get(file_url)
            if file_response.status_code == 200:
                with open(filename, 'wb') as file:
                    file.write(file_response.content)
                print(f"Downloaded: {filename}")

            else:
                print(f"Failed to download: {filename}")
        else:
            print("Skipped a file with a None name.")
else:
    print(f"Failed to retrieve JSON data. Status code: {response.status_code}")     
    
    