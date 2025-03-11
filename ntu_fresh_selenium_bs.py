from bs4 import BeautifulSoup
import os
import re
import time
import argparse
import sys
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# Default credentials (will be overridden by environment variables or command-line arguments)
DEFAULT_USERNAME = "" #ADD EMAIL HERE
DEFAULT_PASSWORD = "" #ADD PASSWORD HERE

def extract_sources_from_html_content(html_content):
    """
    Extract source information from HTML content and return a list of sources.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Find all source spans
    source_spans = soup.find_all('span', id=lambda x: x and x.startswith('lblSource'))
    
    sources = []
    for span in source_spans:
        sources.append(span.text.strip())
    
    return sources

def append_sources_to_file(sources, output_file, existing_sources=None):
    """
    Append sources to a text file, including duplicates.
    
    Args:
        sources (list): List of source strings to append
        output_file (str): Path to the output file
        existing_sources (list, optional): List of all sources (including duplicates)
    
    Returns:
        list: Updated list of all sources (including duplicates)
    """
    # Initialize existing_sources if not provided
    if existing_sources is None:
        existing_sources = []
        # If the file exists, read existing sources
        if os.path.exists(output_file):
            try:
                with open(output_file, 'r', encoding='utf-8') as file:
                    content = file.read()
                    # Extract sources from the numbered list format
                    for line in content.split('\n\n'):
                        if line.strip() and '. ' in line:
                            source = line.split('. ', 1)[1].strip()
                            existing_sources.append(source)
            except Exception as e:
                print(f"Error reading existing sources: {e}")
    
    # Add all sources, including duplicates
    with open(output_file, 'a', encoding='utf-8') as file:
        for source in sources:
            # If the file is not empty and doesn't end with a newline, add one
            if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
                file_content = open(output_file, 'r', encoding='utf-8').read()
                if not file_content.endswith('\n\n'):
                    file.write('\n\n')
            
            # Write the source with its number
            file.write(f"{len(existing_sources) + 1}. {source}\n\n")
            existing_sources.append(source)
    
    print(f"Added {len(sources)} sources to {output_file}")
    return existing_sources

def extract_sources_from_html_file(html_file, output_file, existing_sources=None):
    """
    Extract source information from an HTML file and append to the output file.
    
    Args:
        html_file (str): Path to the HTML file
        output_file (str): Path to the output file
        existing_sources (list, optional): List of existing sources to avoid duplicates
    
    Returns:
        list: Updated list of all sources (including existing ones)
    """
    try:
        with open(html_file, 'r', encoding='utf-8') as file:
            html_content = file.read()
        
        sources = extract_sources_from_html_content(html_content)
        print(f"Found {len(sources)} sources in {html_file}")
        
        return append_sources_to_file(sources, output_file, existing_sources)
    except Exception as e:
        print(f"Error extracting sources from {html_file}: {e}")
        return existing_sources if existing_sources is not None else []

def extract_and_save_sources(output_file='combase_sources.txt'):
    """
    Extract sources from all saved HTML files and save them to a file.
    
    Args:
        output_file (str): Path to the output file
    
    Returns:
        list: List of all sources (including duplicates)
    """
    # Define the HTML files to process
    html_files = [
        'combase_page_1.html',
        'combase_page_2.html',
        'combase_page_3.html',
        'combase_page_4.html',
        'combase_search_results.html'  # Also check the search results page
    ]
    
    # Create or clear the output file
    with open(output_file, 'w', encoding='utf-8') as file:
        pass
    
    all_sources = []
    total_sources = 0
    
    # Process each HTML file
    for html_file in html_files:
        if os.path.exists(html_file):
            print(f"Processing {html_file}...")
            sources_before = len(all_sources)
            all_sources = extract_sources_from_html_file(html_file, output_file, all_sources)
            sources_added = len(all_sources) - sources_before
            total_sources += sources_added
        else:
            print(f"File {html_file} not found.")
    
    print(f"Extracted {total_sources} total sources and saved to {output_file}")
    return all_sources

def login_to_combase(username, password, wait_time=5, headless=False, output_file='combase_sources.txt'):
    """
    Logs into the ComBase Browser website using Selenium and BeautifulSoup.
    
    Args:
        username (str): The username for login
        password (str): The password for login
        wait_time (int): How long to wait between requests (in seconds)
        headless (bool): Whether to run the browser in headless mode
        output_file (str): Path to the output file for sources
    
    Returns:
        webdriver.Chrome: The browser instance if successful, None otherwise
    """
    print(f"Attempting to log in as {username}...")
    
    # Set up Chrome options
    chrome_options = Options()
    if headless:
        print("Running in headless mode")
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    
    # Set up Chrome service with automatic ChromeDriver management
    service = Service(ChromeDriverManager().install())

    # Initialize the Chrome driver
    print("Starting Chrome browser...")
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    # Initialize the list of existing sources
    existing_sources = []
    if os.path.exists(output_file):
        try:
            with open(output_file, 'r', encoding='utf-8') as file:
                content = file.read()
                # Extract sources from the numbered list format
                for line in content.split('\n\n'):
                    if line.strip() and '. ' in line:
                        source = line.split('. ', 1)[1].strip()
                        existing_sources.append(source)
        except Exception as e:
            print(f"Error reading existing sources: {e}")
    
    try:
        # Navigate to the login page
        print("Navigating to ComBase Browser login page...")
        driver.get("https://combasebrowser.errc.ars.usda.gov/membership/Login.aspx?ReturnUrl=%2f")
        
        # Wait for the login form to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "Login1_UserName"))
        )
        
        # Enter login credentials
        print("Entering login credentials...")
        driver.find_element(By.ID, "Login1_UserName").send_keys(username)
        driver.find_element(By.ID, "Login1_Password").send_keys(password)
        
        # Click the login button
        print("Clicking login button...")
        driver.find_element(By.ID, "Login1_Button1").click()
        
        # Wait for redirection after login
        time.sleep(3)
        
        # Check if login was successful
        if "Login.aspx" not in driver.current_url:
            print("Login successful! Redirected to:", driver.current_url)
            
            # Save the login page HTML for debugging
            if driver.page_source is not None:
                with open("combase_login_success.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                print("Login page HTML saved to combase_login_success.html")
            else:
                print("Warning: Page source is None, cannot save login success HTML")
            
            # Now search for "salmonella spp"
            print("Searching for 'salmonella spp'...")
            
            # Try to find and click on the Browser link in the sidebar
            print("Looking for Browser link in the sidebar...")
            
            # Wait for the page to fully load
            time.sleep(2)
            
            # Save the home page HTML for debugging
            if driver.page_source is not None:
                with open("combase_home_page.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                print("Home page HTML saved to combase_home_page.html")
            else:
                print("Warning: Page source is None, cannot save home page HTML")
            
            # Try multiple approaches to find the Browser link
            browser_link = None
            
            # Try by text content
            try:
                browser_link = driver.find_element(By.XPATH, "//a[contains(text(), 'Browser')]")
                print("Found Browser link by text content")
            except NoSuchElementException:
                print("Browser link not found by text content, trying other methods...")
                
                # Try by partial link text
                try:
                    browser_link = driver.find_element(By.PARTIAL_LINK_TEXT, "Browser")
                    print("Found Browser link by partial link text")
                except NoSuchElementException:
                    print("Browser link not found by partial link text, trying other methods...")
                    
                    # Try by href attribute
                    try:
                        browser_link = driver.find_element(By.XPATH, "//a[contains(@href, 'Search.aspx')]")
                        print("Found Browser link by href attribute")
                    except NoSuchElementException:
                        print("Browser link not found by href attribute")
            
            if browser_link:
                # Click the Browser link
                print("Clicking Browser link...")
                try:
                    browser_link.click()
                except Exception as click_error:
                    print(f"Regular click failed: {click_error}")
                    print("Trying JavaScript click...")
                    driver.execute_script("arguments[0].click();", browser_link)
                
                # Wait for the search page to load
                print("Waiting for search page to load...")
                time.sleep(3)
                
                # Save the search page HTML for debugging
                if driver.page_source is not None:
                    with open("combase_search_page.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    print("Search page HTML saved to combase_search_page.html")
                else:
                    print("Warning: Page source is None, cannot save search page HTML")
                
                # Print the current URL
                print("Current URL after clicking Browser link:", driver.current_url)
            else:
                print("ERROR: Could not find Browser link using any method")
                print("Please check the website structure or try again later")
                driver.quit()
                return None
            
            # Save the search page HTML for debugging
            if driver.page_source is not None:
                with open("combase_search_page.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                print("Search page HTML saved to combase_search_page.html")
            else:
                print("Warning: Page source is None, cannot save search page HTML")
            
            # Enter the search term
            print("Entering search term...")
            try:
                # Try to find the search input field with a more specific selector
                search_input = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "div.ms-sel-ctn input"))
                )
                
                print("Found search input field, clicking on it...")
                # Click on the search input field to activate it
                search_input.click()
                
                # Wait a bit after clicking
                time.sleep(1)
                
                print("Typing 'salmonella spp'...")
                # Type "salmonella spp"
                search_input.send_keys("salmonella spp")
                
                # Wait for the dropdown to appear
                print("Waiting for dropdown to appear...")
                time.sleep(1.5)
                
                try:
                    print("Looking for Salmonella spp in dropdown...")
                    
                    # Try multiple selector strategies
                    selectors = [
                        # Simple class-based selector
                        "div.ms-res-item",
                        # More specific selector
                        "div.ms-res-item.ms-res-item-active",
                        # First item in dropdown
                        ".ms-res-ctn .ms-res-item:first-child"
                    ]
                    
                    dropdown_item = None
                    for selector in selectors:
                        try:
                            print(f"Trying selector: {selector}")
                            items = driver.find_elements(By.CSS_SELECTOR, selector)
                            print(f"Found {len(items)} items with selector {selector}")
                            
                            for item in items:
                                item_text = item.text.lower()
                                print(f"Item text: {item_text}")
                                if "salmonella spp" in item_text:
                                    dropdown_item = item
                                    print(f"Found matching item: {item_text}")
                                    break
                            
                            if dropdown_item:
                                break
                        except Exception as e:
                            print(f"Error with selector {selector}: {e}")
                    
                    if dropdown_item:
                        print("Clicking on Salmonella spp option...")
                        # Try regular click first
                        try:
                            dropdown_item.click()
                        except Exception as click_error:
                            print(f"Regular click failed: {click_error}")
                            print("Trying JavaScript click...")
                            driver.execute_script("arguments[0].click();", dropdown_item)
                        
                        # Wait a bit after selection
                        time.sleep(1)
                        print("Option selected successfully")
                    else:
                        print("Could not find Salmonella spp in dropdown")
                        print("Continuing with search anyway...")
                except Exception as dropdown_error:
                    print(f"Error selecting from dropdown: {dropdown_error}")
                    print("Continuing with search anyway...")
            except Exception as e:
                print(f"Error finding search input field: {e}")
                # Try alternative methods if the specific selector fails
                try:
                    # Try by ID
                    organism_input = driver.find_element(By.ID, "ContentPlaceHolder1_txtOrganism")
                    print("Found organism input field with ID: ContentPlaceHolder1_txtOrganism")
                    organism_input.clear()
                    organism_input.send_keys("salmonella spp")
                except NoSuchElementException:
                    print("Organism input field with ID 'ContentPlaceHolder1_txtOrganism' not found, trying other methods...")
                    
                    # Try by name attribute
                    try:
                        organism_inputs = driver.find_elements(By.XPATH, "//input[@name='ctl00$ContentPlaceHolder1$txtOrganism']")
                        if organism_inputs:
                            print(f"Found organism input field by name")
                            organism_inputs[0].clear()
                            organism_inputs[0].send_keys("salmonella spp")
                        else:
                            print("No input field found by name")
                    except Exception as input_error:
                        print(f"Error finding input field by name: {input_error}")
            
            # Click the search button
            print("Clicking search button...")
            try:
                # Find the search button with ID
                search_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "btnDoSearch"))
                )
                print("Found search button with ID: btnDoSearch")
                search_button.click()
            except Exception as search_error:
                print(f"Error finding search button with ID 'btnDoSearch': {search_error}")
                
                # Try alternative methods
                try:
                    # Try with ContentPlaceHolder prefix
                    search_button = driver.find_element(By.ID, "ContentPlaceHolder1_btnDoSearch")
                    print("Found search button with ID: ContentPlaceHolder1_btnDoSearch")
                    search_button.click()
                except NoSuchElementException:
                    print("Search button with ID 'ContentPlaceHolder1_btnDoSearch' not found, trying other methods...")
                    
                    # Try by value
                    try:
                        search_buttons = driver.find_elements(By.XPATH, "//input[@type='submit' and @value='Search']")
                        if search_buttons:
                            print(f"Found {len(search_buttons)} search buttons with value 'Search'")
                            search_buttons[0].click()
                            print("Clicked the first search button found")
                        else:
                            print("No search buttons found with value 'Search'")
                    except Exception as e:
                        print(f"Error finding search buttons by value: {e}")
            
            # Wait for the search results to load
            print("Waiting for search results to load...")
            time.sleep(3)
            
            # Check if we were redirected to the search results page
            print("Current URL after search:", driver.current_url)
            
            if "SearchResults.aspx" in driver.current_url:
                print("Successfully redirected to search results page")
                
                # Save the search results HTML for debugging
                if driver.page_source is not None:
                    with open("combase_search_results.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    print("Search results HTML saved to combase_search_results.html")
                    
                    # Extract sources from the search results page
                    sources = extract_sources_from_html_content(driver.page_source)
                    print(f"Found {len(sources)} sources in search results page")
                    existing_sources = append_sources_to_file(sources, output_file, existing_sources)
                else:
                    print("Warning: Page source is None, cannot save search results HTML")
                
                # Parse the search results with BeautifulSoup
                soup = BeautifulSoup(driver.page_source, 'html.parser')
                
                # Try to find the total number of pages
                total_pages = 1  # Default to 1 if we can't find the total
                hidden_total_pages = soup.find('input', {'id': 'HiddenTotalPages'})
                if hidden_total_pages and hidden_total_pages.get('value'):
                    try:
                        total_pages = int(hidden_total_pages['value'])
                        print(f"Total pages of results: {total_pages}")
                    except ValueError:
                        print("Could not determine total pages, assuming 1 page")
                
                # Process each page of results
                current_page = 1
                
                while current_page <= total_pages:
                    print(f"\nProcessing page {current_page} of {total_pages}...")
                    
                    # Save the current page HTML
                    if driver.page_source is not None:
                        with open(f"combase_page_{current_page}.html", "w", encoding="utf-8") as f:
                            f.write(driver.page_source)
                        print(f"Page HTML saved to combase_page_{current_page}.html")
                        
                        # Extract sources from the current page
                        sources = extract_sources_from_html_content(driver.page_source)
                        print(f"Found {len(sources)} sources in page {current_page}")
                        existing_sources = append_sources_to_file(sources, output_file, existing_sources)
                    else:
                        print(f"Warning: Page source is None, cannot save page {current_page} HTML")
                    
                    # Find all checkboxes for export
                    print("Finding export checkboxes...")
                    checkboxes = driver.find_elements(By.CSS_SELECTOR, "input.exportchk")
                    
                    if checkboxes:
                        print(f"Found {len(checkboxes)} checkboxes")
                        
                        # Select all checkboxes
                        for checkbox in checkboxes:
                            if not checkbox.is_selected():
                                print(f"Selecting checkbox {checkbox.get_attribute('id')}...")
                                checkbox.click()
                                time.sleep(0.2)  # Small delay between selections
                        
                        # Click the export button
                        print("Clicking export button...")
                        try:
                            export_button = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, "cbBtnExportToExcel"))
                            )
                            
                            # Scroll to the export button to make it visible
                            print("Scrolling to export button...")
                            driver.execute_script("arguments[0].scrollIntoView(true);", export_button)
                            time.sleep(1)  # Give time for scrolling to complete
                            
                            # Use JavaScript to click the button
                            print(f"Clicking export button for page {current_page}...")
                            driver.execute_script("arguments[0].click();", export_button)
                            
                            # Wait for the export to complete
                            print("Waiting for export to complete...")
                            time.sleep(3)
                            
                            # Take a screenshot after export
                            export_screenshot = f"combase_export_page_{current_page}.png"
                            driver.save_screenshot(export_screenshot)
                            print(f"Export screenshot saved to {export_screenshot}")
                            
                            # Check if the file was downloaded
                            print("Export completed. Check your downloads folder for the Excel file.")
                            
                            # Deselect all checkboxes before moving to the next page
                            print(f"Deselecting all checkboxes on page {current_page}...")
                            for checkbox in checkboxes:
                                if checkbox.is_selected():
                                    print(f"Deselecting checkbox {checkbox.get_attribute('id')}...")
                                    checkbox.click()
                                    time.sleep(0.2)  # Small delay between deselections
                            
                            print(f"Deselected {len(checkboxes)} checkboxes on page {current_page}")
                            
                            # Take a screenshot after deselecting checkboxes
                            deselect_screenshot = f"combase_deselect_page_{current_page}.png"
                            driver.save_screenshot(deselect_screenshot)
                            print(f"Deselect screenshot saved to {deselect_screenshot}")
                        except Exception as export_error:
                            print(f"Error with export button: {export_error}")
                            
                            # Try alternative method
                            try:
                                # Try with ContentPlaceHolder prefix
                                export_button = driver.find_element(By.ID, "ContentPlaceHolder1_cbBtnExportToExcel")
                                print("Found export button with ID: ContentPlaceHolder1_cbBtnExportToExcel")
                                
                                # Scroll to the export button to make it visible
                                driver.execute_script("arguments[0].scrollIntoView(true);", export_button)
                                time.sleep(1)  # Give time for scrolling to complete
                                
                                # Use JavaScript to click the button
                                driver.execute_script("arguments[0].click();", export_button)
                                
                                # Wait for the export to complete
                                print("Waiting for export to complete...")
                                time.sleep(3)
                                
                                print("Export completed. Check your downloads folder for the Excel file.")
                                
                                # Deselect all checkboxes before moving to the next page
                                print(f"Deselecting all checkboxes on page {current_page}...")
                                for checkbox in checkboxes:
                                    if checkbox.is_selected():
                                        print(f"Deselecting checkbox {checkbox.get_attribute('id')}...")
                                        checkbox.click()
                                        time.sleep(0.2)  # Small delay between deselections
                                
                                print(f"Deselected {len(checkboxes)} checkboxes on page {current_page}")
                                
                                # Take a screenshot after deselecting checkboxes
                                deselect_screenshot = f"combase_deselect_page_{current_page}.png"
                                driver.save_screenshot(deselect_screenshot)
                                print(f"Deselect screenshot saved to {deselect_screenshot}")
                            except NoSuchElementException:
                                print("Export button not found with any known ID")
                    else:
                        print("No checkboxes found for export")
                    
                    # If this is not the last page, go to the next page
                    if current_page < total_pages:
                        print(f"Navigating to page {current_page + 1}...")
                        
                        # Find the next button
                        try:
                            next_button = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, "a.next[data-action='next']"))
                            )
                            
                            # Scroll to the next button to make it visible
                            print("Scrolling to next button...")
                            driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                            time.sleep(0.5)  # Give time for scrolling to complete
                            
                            # Use JavaScript to click the button to avoid interception issues
                            print("Clicking next button using JavaScript...")
                            driver.execute_script("arguments[0].click();", next_button)
                            
                            # Wait for the next page to load
                            print("Waiting for next page to load...")
                            time.sleep(2)
                            
                            # Update the soup object with the new page source
                            soup = BeautifulSoup(driver.page_source, 'html.parser')
                        except Exception as next_error:
                            print(f"Error navigating to next page: {next_error}")
                            
                            # Try alternative method
                            try:
                                # Try to find the next page link by page number
                                next_page_link = driver.find_element(By.XPATH, f"//a[contains(text(), '{current_page + 1}')]")
                                print(f"Found next page link for page {current_page + 1}")
                                next_page_link.click()
                                
                                # Wait for the next page to load
                                print("Waiting for next page to load...")
                                time.sleep(2)
                                
                                # Update the soup object with the new page source
                                soup = BeautifulSoup(driver.page_source, 'html.parser')
                            except NoSuchElementException:
                                print(f"Next page link not found")
                                break
                    
                    current_page += 1
                
                print("\nAll pages processed successfully!")
                print(f"Total sources extracted: {len(existing_sources)}")
            else:
                print("Not redirected to search results page")
                
                # Save the current page HTML for debugging
                if driver.page_source is not None:
                    with open("combase_search_failure.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    print("Current page HTML saved to combase_search_failure.html")
                else:
                    print("Warning: Page source is None, cannot save search failure HTML")
            
            return driver
        else:
            # Check if there's an error message
            try:
                error_message = driver.find_element(By.ID, "Login1_FailureText").text
                print(f"Login failed. Error message: {error_message}")
            except NoSuchElementException:
                print("Login verification failed. Still on login page.")
            
            # Save the login failure page HTML for debugging
            if driver.page_source is not None:
                with open("combase_login_failure.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                print("Login failure page HTML saved to combase_login_failure.html")
            else:
                print("Warning: Page source is None, cannot save login failure HTML")
            
            driver.quit()
            return None
    except Exception as e:
        print(f"An error occurred: {e}")
        driver.quit()
        return None

def combine_excel_files(output_file='ComBaseCombined.xlsx'):
    """
    Combines all ComBaseExport.xlsx files in the Downloads directory into a single Excel file with multiple tabs.
    
    Args:
        output_file (str): Path to the output Excel file
    
    Returns:
        bool: True if successful, False otherwise
    """
    # Get the Downloads directory path
    downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
    
    # Find all ComBaseExport.xlsx files in the Downloads directory
    excel_files = glob.glob(os.path.join(downloads_dir, 'ComBaseExport*.xlsx'))
    
    if not excel_files:
        print("No ComBaseExport Excel files found in Downloads directory.")
        return False
    
    print(f"Found {len(excel_files)} Excel files to combine.")
    
    # Initialize dictionaries to store data from each sheet
    data_records = []
    logs = []
    
    # Read each Excel file and extract data from both sheets
    for file in excel_files:
        print(f"Processing {file}...")
        try:
            # Read the Excel file
            xls = pd.ExcelFile(file)
            
            # Get the sheet names
            sheet_names = xls.sheet_names
            
            # Check if the file has at least two sheets
            if len(sheet_names) >= 2:
                # Read the first sheet (data records)
                df_data = pd.read_excel(xls, sheet_name=0)
                data_records.append(df_data)
                
                # Read the second sheet (logs)
                df_logs = pd.read_excel(xls, sheet_name=1)
                logs.append(df_logs)
            else:
                print(f"Warning: {file} does not have at least two sheets. Skipping.")
        except Exception as e:
            print(f"Error processing {file}: {e}")
    
    if not data_records or not logs:
        print("No valid data found in Excel files.")
        return False
    
    # Combine all data records and logs
    combined_data = pd.concat(data_records, ignore_index=True)
    combined_logs = pd.concat(logs, ignore_index=True)
    
    # Create a new Excel file with multiple sheets
    output_path = os.path.join(downloads_dir, output_file)
    
    print(f"Creating combined Excel file at {output_path}...")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        combined_data.to_excel(writer, sheet_name='Data Records', index=False)
        combined_logs.to_excel(writer, sheet_name='Logs', index=False)
    
    print(f"Successfully combined {len(excel_files)} Excel files into {output_path}")
    return True

def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description='Login to ComBase Browser using Selenium and BeautifulSoup')
    
    parser.add_argument('-u', '--username', 
                        help='Username for ComBase login')
    
    parser.add_argument('-p', '--password', 
                        help='Password for ComBase login')
    
    parser.add_argument('-w', '--wait', type=int, default=5,
                        help='How long to wait between requests (in seconds)')
    
    parser.add_argument('--headless', action='store_true',
                        help='Run the browser in headless mode')
    
    parser.add_argument('--extract-only', action='store_true',
                        help='Only extract sources from existing HTML files without running Selenium')
    
    parser.add_argument('-o', '--output', default='combase_sources.txt',
                        help='Output file for sources')
    
    parser.add_argument('--combine-excel', action='store_true',
                        help='Combine all ComBaseExport Excel files in Downloads directory')
    
    parser.add_argument('--excel-output', default='ComBaseCombined.xlsx',
                        help='Output file for combined Excel data')
    
    return parser.parse_args()

if __name__ == "__main__":
    # Parse command line arguments
    args = parse_arguments()
    
    # Check if we should only combine Excel files
    if args.combine_excel:
        print("Combining Excel files...")
        if combine_excel_files(args.excel_output):
            print("Excel files combined successfully.")
        else:
            print("Failed to combine Excel files.")
        sys.exit(0)
    
    # Check if we should only extract sources from existing HTML files
    if args.extract_only:
        print("Extracting sources from existing HTML files...")
        extract_and_save_sources(args.output)
        sys.exit(0)
    
    # Get credentials from environment variables or command line arguments or use defaults
    username = args.username or os.environ.get('COMBASE_USERNAME') or DEFAULT_USERNAME
    password = args.password or os.environ.get('COMBASE_PASSWORD') or DEFAULT_PASSWORD
    
    # Check if credentials are provided
    if not username or not password:
        print("Error: Username and password are required.")
        print("Provide them via command line arguments (-u, -p) or environment variables (COMBASE_USERNAME, COMBASE_PASSWORD)")
        sys.exit(1)
    
    # Login to ComBase
    driver = login_to_combase(username, password, wait_time=args.wait, headless=args.headless, output_file=args.output)
    
    if driver:
        print("Script completed successfully")
        driver.quit()
        
        # Ask if the user wants to combine Excel files
        combine_files = input("Do you want to combine exported Excel files? (y/n): ").strip().lower()
        if combine_files == 'y' or combine_files == 'yes':
            if combine_excel_files(args.excel_output):
                print("Excel files combined successfully.")
            else:
                print("Failed to combine Excel files.")
    else:
        print("Script failed")
        sys.exit(1)
