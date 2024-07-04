import requests
from bs4 import BeautifulSoup
import csv
import time
import os
import logging
import random

# Set up logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def get_random_user_agent():
    user_agents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0'
    ]
    return random.choice(user_agents)

def scrape_launch_data(session, url):
    logging.info(f"Attempting to scrape data from {url}")
    try:
        headers = {
            'User-Agent': get_random_user_agent(),
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Referer': 'https://spacelaunchnow.me/',
            'DNT': '1',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        }
        response = session.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        logging.debug(f"Successfully retrieved page: {url}")
    except requests.RequestException as e:
        logging.error(f"Failed to retrieve {url}: {e}")
        return []

    soup = BeautifulSoup(response.text, 'html.parser')
    launches = []
    table = soup.find('table', class_='table')
    
    if table:
        logging.debug("Found table on the page")
        rows = table.find_all('tr')[1:]  # Skip header row
        logging.info(f"Found {len(rows)} rows in the table")
        for row in rows:
            cols = row.find_all('td')
            if len(cols) == 7:
                launch = {
                    'name': cols[0].text.strip(),
                    'status': cols[1].text.strip(),
                    'provider': cols[2].text.strip(),
                    'rocket': cols[3].text.strip(),
                    'mission': cols[4].text.strip(),
                    'date': cols[5].text.strip(),
                    'pad': cols[6].text.strip()
                }
                launches.append(launch)
                logging.debug(f"Extracted launch: {launch['name']}")
    else:
        logging.warning("No table found on the page")
    
    logging.info(f"Scraped {len(launches)} launches from the page")
    return launches

def save_to_csv(launches, filename, mode='a'):
    file_exists = os.path.isfile(filename)
    logging.info(f"Saving {len(launches)} launches to {filename} in {mode} mode")
    
    try:
        with open(filename, mode, newline='', encoding='utf-8') as csvfile:
            fieldnames = ['name', 'status', 'provider', 'rocket', 'mission', 'date', 'pad']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            if not file_exists or mode == 'w':
                writer.writeheader()
                logging.debug("Wrote header to CSV file")
            
            for launch in launches:
                writer.writerow(launch)
            
        logging.info(f"Successfully saved data to {filename}")
    except IOError as e:
        logging.error(f"Error writing to CSV file: {e}")

def main():
    base_url = "https://spacelaunchnow.me/launch/"
    output_file = 'space_launches.csv'
    
    session = requests.Session()
    
    # Start from the last scraped page or page 1 if starting fresh
    try:
        with open('last_page.txt', 'r') as f:
            start_page = int(f.read().strip()) + 1
        logging.info(f"Resuming from page {start_page}")
    except FileNotFoundError:
        start_page = 1
        logging.info("Starting fresh from page 1")
    
    for page in range(start_page, 291):  # Pages start_page to 290
        url = f"{base_url}?page={page}"
        logging.info(f"Processing page {page}")
        
        try:
            launches = scrape_launch_data(session, url)
            
            if launches:
                # Save mode: 'w' for first page, 'a' for subsequent pages
                mode = 'w' if page == start_page else 'a'
                save_to_csv(launches, output_file, mode)
                
                # Update last scraped page
                with open('last_page.txt', 'w') as f:
                    f.write(str(page))
                logging.debug(f"Updated last_page.txt with {page}")
            else:
                logging.warning(f"No launches found on page {page}")
            
            # Randomized delay between requests
            time.sleep(random.uniform(2, 5))
        
        except Exception as e:
            logging.exception(f"An unexpected error occurred while processing page {page}")
            break
    
    logging.info("Scraping complete.")

if __name__ == "__main__":
    main()