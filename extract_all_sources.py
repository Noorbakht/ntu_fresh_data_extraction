from bs4 import BeautifulSoup
import os

def extract_sources_from_html_file(html_file):
    """
    Extract source information from an HTML file and return a list of sources.
    """
    try:
        with open(html_file, 'r', encoding='utf-8') as file:
            html_content = file.read()
        
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Find all source spans
        source_spans = soup.find_all('span', id=lambda x: x and x.startswith('lblSource'))
        
        sources = []
        for span in source_spans:
            sources.append(span.text.strip())
        
        return sources
    except Exception as e:
        print(f"Error reading file {html_file}: {e}")
        return []

def save_sources_to_file(sources, output_file):
    """
    Save the list of sources to a text file.
    """
    with open(output_file, 'w', encoding='utf-8') as file:
        for i, source in enumerate(sources, 1):
            file.write(f"{i}. {source}\n\n")

def main():
    # Define the HTML files to process
    html_files = [
        'combase_search_results.html',
        'combase_page_1.html',
        'combase_page_2.html',
        'combase_page_3.html',
        'combase_page_4.html'
    ]
    
    all_sources = []
    
    # Process each HTML file
    for html_file in html_files:
        if os.path.exists(html_file):
            print(f"Processing {html_file}...")
            sources = extract_sources_from_html_file(html_file)
            print(f"Found {len(sources)} sources in {html_file}")
            all_sources.extend(sources)
        else:
            print(f"File {html_file} not found.")
    
    # Remove duplicates while preserving order
    unique_sources = []
    for source in all_sources:
        if source not in unique_sources:
            unique_sources.append(source)
    
    # Save all sources to a file
    output_file = 'all_combase_sources.txt'
    save_sources_to_file(unique_sources, output_file)
    print(f"Extracted {len(unique_sources)} unique sources and saved to {output_file}")
    
    # Print the first few sources
    if unique_sources:
        print("\nFirst few sources:")
        for i, source in enumerate(unique_sources[:3], 1):
            print(f"{i}. {source}")

if __name__ == "__main__":
    main()
