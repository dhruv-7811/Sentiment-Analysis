import pandas as pd
import requests
from bs4 import BeautifulSoup
import os


def extract_article_text(url):                          # Function to extract article text from a URL
    try:
        response = requests.get(url)
        response.raise_for_status()  
        soup = BeautifulSoup(response.text, 'html.parser')
        
        article_div = soup.find('div', class_='tdb-block-inner td-fix-index')
        if article_div:
            text = article_div.get_text(separator='\n')
            return text.strip()
        else:
            print(f"Div with class 'td-post-content tagdiv-type' not found in {url}")
            return ""
    except Exception as e:
        print(f"Failed to extract text from {url}: {e}")
        return ""
def save_text_to_file(text, filename):
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(text)

def main(excel_file):
    # Read the Excel file
    df = pd.read_excel(excel_file)
    
    
    # Create a directory to save the text files
    output_dir = 'articles'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Process each URL
    for index, row in df.iterrows():
        url = row['URL']
        if pd.isna(url):
            continue
        
        print(f"Processing URL: {url}")
        article_text = extract_article_text(url)
        
        if article_text:
            # Create a valid filename from the URL
            filename = os.path.join(output_dir, f"{row['URL_ID']}.txt")
            save_text_to_file(article_text, filename)
            print(f"Saved article to {filename}")
        else:
            print(f"Failed to extract article from {url}")

#main funtion
if __name__ == "__main__":
    excel_file = 'Input.xlsx'  
    main(excel_file)
