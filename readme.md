# Space Launch Data Scraper

This project contains a Python script (`scraper.py`) designed to scrape space launch data from the Space Launch Now website and save it to a CSV file. The script is capable of resuming from the last scraped page, ensuring that no data is missed even if the script is interrupted.

## Features

- **Random User-Agent**: Uses a random User-Agent for each request to avoid being blocked by the website.
- **Logging**: Comprehensive logging to track the progress and any issues encountered during the scraping process.
- **CSV Output**: Saves the scraped data to a CSV file, appending new data if the file already exists.
- **Resume Capability**: Automatically resumes from the last scraped page, as recorded in `last_page.txt`.
- **Randomized Delays**: Introduces random delays between requests to mimic human browsing behavior.

## Requirements

- Python 3.x
- `requests` library
- `beautifulsoup4` library

You can install the required libraries using pip:
