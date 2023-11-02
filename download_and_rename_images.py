import os
import openpyxl
import requests
from urllib.parse import urlparse
import time

# Function to download an image from a website URL
def download_image_from_url(url, destination):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        with open(destination, 'wb') as f:
            f.write(response.content)
        return True
    except requests.exceptions.RequestException as e:
        print(f"Failed to download image from {url}: {e}")
        return False

# Load the Excel file
excel_file = openpyxl.load_workbook('ChancellorsAwards2023Submissions.xlsx')
sheet = excel_file['ChancellorsAwards2023Submission']  # Change to the appropriate sheet name

# Define the column with the names and the column with the website URLs
name_col = 10  # Column A
url_col = 16   # Column B

# Loop through rows in the Excel file
for row in sheet.iter_rows(min_row=2, values_only=True):  # Assuming headers in row 1
    name = row[name_col - 1]
    url = row[url_col - 1]

    # Check if the URL is valid
    if not url:
        continue

    # Parse the URL to get the filename
    filename = os.path.basename(urlparse(url).path)

    # Download the image from the website URL
    if download_image_from_url(url, filename):
        # Rename the file to the name
        new_filename = f"{name}{os.path.splitext(filename)[1]}"
        os.rename(filename, new_filename)
    
    # Add a delay to respect rate limiting
    time.sleep(1)  # Wait for 1 second between requests

# Close the Excel file
excel_file.close()
