# NSE Shareholding Patterns Scraper

This repository contains a Python script that scrapes the National Stock Exchange (NSE) website and extracts the shareholding patterns of any listed company for all financial quarters, going back to 2016. The extracted data includes the Summary Statement holding of specified securities and the Statement showing shareholding pattern of the Public shareholder. The data is organized into properly formatted Excel files.

## Features

- Scrapes shareholding patterns from the NSE website.
- Extracts data for all financial quarters from 2016 onwards.
- Organizes data into well-formatted Excel files.
- Downloads shareholding patterns of a quarter in under 30 seconds.
- Handles exceptions for companies that cannot be scraped (to be improved in future updates).

## Requirements

- Python 3.10+
- `selenium`
- `openpyxl`
- `webdriver_manager`

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/NSE-Shareholding-Patterns-Scraper.git
    cd NSE-Shareholding-Patterns-Scraper
    ```

2. Install the required packages:
    ```sh
    pip install -r requirements.txt
    ```

## Usage

1. Open the main.py file and modify the `initiate` function call at the bottom to specify the company name you want to scrape:
    ```python
    if __name__ == "__main__":
        initiate('CIPLA')
    ```

2. Run the script:
    ```sh
    python main.py
    ```

3. The script will create a folder named Shareholding_Patterns and save the Excel files with the shareholding patterns of the specified company.

## File Structure

- main.py: The main script that initiates the scraping process.
- data_scraper.py: Contains the function to scrape data from the NSE website.
- excel_formatter.py: Contains the function to format the Excel worksheets.
- selenium_utils.py: Contains utility functions for setting up the Selenium WebDriver and scrolling the webpage.
- Shareholding_Patterns: Folder where the Excel files are saved.

## Future Improvements

- Handle exceptions for companies that cannot be scraped.
- Optimize the scraping process for faster execution.
- Add more detailed error handling and logging.

---

Feel free to reach out if you have any questions or suggestions!
