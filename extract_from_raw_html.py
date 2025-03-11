from bs4 import BeautifulSoup
import re

def main():
    # Read the raw HTML content from the file
    try:
        with open('combase_search_results.html', 'r', encoding='utf-8') as file:
            html_content = file.read()
        
        # Parse the HTML content
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Find all source spans using a more specific selector
        source_spans = soup.find_all('span', id=re.compile(r'^lblSource$'))
        
        if not source_spans:
            # Try alternative selectors if the first one doesn't work
            source_spans = soup.find_all('span', class_='text-primary', string='Source')
            if source_spans:
                # If we found the "Source" labels, get their parent divs and then find the next div with the actual source text
                sources = []
                for span in source_spans:
                    parent_div = span.find_parent('div')
                    if parent_div and parent_div.find_next_sibling('div'):
                        source_div = parent_div.find_next_sibling('div')
                        source_span = source_div.find('span')
                        if source_span:
                            sources.append(source_span.text.strip())
            else:
                # Try another approach - find all divs with source information
                sources = []
                rows = soup.find_all('div', class_='cbRowSummaryResult')
                for row in rows:
                    source_div = row.find('div', string=lambda s: s and 'Source' in s)
                    if source_div and source_div.find_next_sibling('div'):
                        source_text_div = source_div.find_next_sibling('div')
                        source_span = source_text_div.find('span')
                        if source_span:
                            sources.append(source_span.text.strip())
        else:
            # Extract the text from each source span
            sources = [span.text.strip() for span in source_spans]
        
        print(f"Found {len(sources)} sources in the HTML content")
        
        # Save sources to a file
        output_file = 'sources_from_raw_html.txt'
        with open(output_file, 'w', encoding='utf-8') as file:
            for i, source in enumerate(sources, 1):
                file.write(f"{i}. {source}\n\n")
        
        print(f"Saved {len(sources)} sources to {output_file}")
        
        # Print all sources
        print("\nAll sources:")
        for i, source in enumerate(sources, 1):
            print(f"{i}. {source}")
    
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
