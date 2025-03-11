# NTU Fresh BeautifulSoup Data Extraction Tool

## What This Tool Does

This collection of scripts helps you extract data from ComBase Browser using Selenium and BeautifulSoup, and then process the extracted data. The main workflow involves running scripts in sequence to:

1. Log into ComBase Browser automatically
2. Search for "salmonella spp"
3. Extract source information from search results
4. Export data to Excel files in your Downloads folder
5. Combine all exported Excel files into a single file

## Prerequisites

Before using this tool, make sure you have the following:

### Software Requirements

- Python 3.x installed
- Chrome browser installed
- Internet connection

### Python Libraries

Install the required Python libraries using pip:

```
pip install selenium beautifulsoup4 pandas openpyxl webdriver-manager
wget https://dl.google.com/linux/direct/google-chrome-stable_current_x86_64.rpm
sudo yum install -y google-chrome-stable_current_x86_64.rpm
sudo curl https://intoli.com/install-google-chrome.sh | bash
wget https://chromedriver.storage.googleapis.com/[version]/chromedriver_linux64.zip
unzip chromedriver_linux64.zip
sudo mv chromedriver /usr/bin/chromedriver
sudo chown root:root /usr/bin/chromedriver
sudo chmod +x /usr/bin/chromedriver
```

### ComBase Account

- Valid ComBase Browser account (username and password)
- Access permissions to search and export data

## What are Selenium and BeautifulSoup?

### Selenium

Selenium is a powerful automation tool for web browsers. It allows you to:

- Control a web browser programmatically
- Automate interactions with websites (clicking buttons, filling forms, etc.)
- Navigate between web pages
- Wait for specific elements to load
- Handle login processes and other complex interactions

In this project, Selenium is used to automate the process of logging into ComBase Browser, searching for "salmonella spp", and exporting data to Excel files.

### BeautifulSoup

BeautifulSoup is a Python library for parsing HTML and XML documents. It helps you:

- Extract data from HTML pages
- Navigate and search the HTML structure
- Find specific elements using various selectors
- Extract text and attributes from HTML elements

In this project, BeautifulSoup is used to parse the HTML content of ComBase Browser pages and extract source information from specific elements.

## Scripts Explained

### 1. ntu_fresh_selenium_bs.py

The main script that automates the entire process:

- Uses Selenium to control a Chrome browser
- Logs into ComBase Browser with provided credentials
- Searches for "salmonella spp"
- Extracts source information from search results
- Exports data to Excel files in your Downloads folder
- Includes a function to combine Excel files
- Can be run with various command-line options

### 2. extract_all_sources.py

A utility script for extracting sources from saved HTML files:

- Processes multiple HTML files saved by ntu_fresh_selenium_bs.py
- Works with files like combase_search_results.html, combase_page_1.html, etc.
- Extracts source information using BeautifulSoup
- Saves unique sources to a text file (all_combase_sources.txt)

### 3. extract_from_raw_html.py

A simpler script for extracting sources from a single HTML file:

- Processes only combase_search_results.html
- Uses multiple approaches to find source information in the HTML
- Saves sources to a text file (sources_from_raw_html.txt)

### 4. test_combine_excel.py

A simple script for combining Excel files:

- Calls the combine_excel_files() function from ntu_fresh_selenium_bs.py
- Finds all ComBaseExport\*.xlsx files in your Downloads folder
- Combines them into a single Excel file (ComBaseCombined.xlsx)
- Creates two tabs: Data Records and Logs

## How to Run the Tool

### Step 1: Extract Data with Selenium and BeautifulSoup

Run the main script with an output file parameter:

```
python3 ntu_fresh_selenium_bs.py -o output_file.txt
```

This will:

- Launch a Chrome browser
- Log into ComBase Browser
- Search for "salmonella spp"
- Extract source information
- Save sources to the specified output file
- Export data to Excel files in your Downloads folder

### Step 2: Combine Excel Files

After running the main script, run the test_combine_excel.py script:

```
python3 test_combine_excel.py
```

This will:

- Find all ComBaseExport\*.xlsx files in your Downloads folder
- Combine them into a single Excel file named "ComBaseCombined.xlsx"
- Create two tabs in the combined file:
  1. **Data Records** - Contains all your data in one place
  2. **Logs** - Contains all the log information in one place

## Command Line Options

### Main Script (ntu_fresh_selenium_bs.py)

```
python3 ntu_fresh_selenium_bs.py -o output_file.txt [options]
```

Options:

- `-o, --output`: Specify the output file for sources (default: combase_sources.txt)
- `-u, --username`: Username for ComBase login
- `-p, --password`: Password for ComBase login
- `-w, --wait`: How long to wait between requests in seconds (default: 5)
- `--headless`: Run the browser in headless mode
- `--extract-only`: Only extract sources from existing HTML files without running Selenium
- `--combine-excel`: Combine all ComBaseExport Excel files in Downloads directory
- `--excel-output`: Output file for combined Excel data (default: ComBaseCombined.xlsx)

### Combine Excel Script (test_combine_excel.py)

```
python3 test_combine_excel.py
```

This script has no command line options. It will always create a file named "ComBaseCombined.xlsx" in your Downloads folder.

## Where to Find Your Files

- **Source Text File**: The location you specified with the `-o` parameter
- **Combined Excel File**: Your Downloads folder with the name "ComBaseCombined.xlsx"

## Troubleshooting

If you encounter issues:

1. Make sure you have the required libraries installed:

   ```
   pip install selenium beautifulsoup4 pandas openpyxl webdriver-manager
   ```

2. Check that you have Chrome browser installed

3. For Excel combination issues, make sure:
   - You have Excel files in your Downloads folder that start with "ComBaseExport"
   - Each file has at least two sheets (tabs)
