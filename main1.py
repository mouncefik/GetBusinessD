import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import os

# Function to extract emails from website URL
def extract_emails_from_website(url):
    try:
        if not url.startswith('http://') and not url.startswith('https://'):
            url = 'https://' + url  # Add 'https://' if protocol is missing
        response = requests.get(url)
        response.raise_for_status()  # Raise an exception for invalid response codes
        html_content = response.text
        soup = BeautifulSoup(html_content, 'html.parser')
        emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b', soup.get_text())
        return emails
    except Exception as e:
        print(f"Error occurred while scraping {url}: {e}")
        return []

def main():
    excel_file = 'google_maps_data_yoga+Maroc.xlsx'
    output_file = 'emails_from_websites3.xlsx'
    try:
        df = pd.read_excel(excel_file)
        if 'Website' not in df.columns:
            print("Error: 'Website' column not found in the Excel file.")
            return

        all_emails = []
        for index, row in df.iterrows():
            if 'Website' in row:
                url = row['Website']
                print(f"Scraping emails from {url}")
                emails = extract_emails_from_website(url)
                all_emails.extend(emails)
            else:
                print("Warning: Row does not contain a 'Website' column. Skipping.")

        # Remove duplicates from the list of emails
        unique_emails = list(set(all_emails))

        df_emails = pd.DataFrame({'Emails': unique_emails})

        output_path = os.path.join(os.getcwd(), output_file)
        df_emails.to_excel(output_path, index=False)
        print(f"Emails scraped from websites have been saved to {output_path}")
    except FileNotFoundError:
        print(f"Error: File '{excel_file}' not found.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
