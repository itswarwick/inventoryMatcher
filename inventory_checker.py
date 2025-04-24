import os
import pandas as pd
import pdfplumber
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
import argparse

class InventoryChecker:
    def __init__(self):
        self.base_url = "https://southernstarz.com"
        self.driver = None
        self.producers = [
            'Mount Fishtail',
            'Sherwood',
            'Stratum',
            'Black Pearl Vineyards',
            'Downes Family Vineyards',
            'David Finlayson',
            'Luddite',
            'Painted Wolf',
            'Thistledown',
            'Nugan Estate',
            'RL Buller',
            'Water Wheel',
            'Casas del Bosque',
            'Earthsong',
            'Miha',
            'Almarada',
            'Antucura',
            'Vina Alicia',
            'Greenock Creek',
            'Oliverhill',
            'Tait',
            'Wildberry Estate',
            'NEWBLOOD',
            'Rieslingfreak'
        ]
        # Also store shortened versions for matching
        self.producer_variants = {
            'BLACK PEARL': 'Black Pearl Vineyards',
            'DOWNES FAMILY': 'Downes Family Vineyards',
            'CASAS DEL BOSQUE': 'Casas del Bosque',
            'WILDBERRY': 'Wildberry Estate',
            'VINA ALICIA': 'Vina Alicia',
            'NUGAN': 'Nugan Estate',
            'GREENOCK': 'Greenock Creek'
        }
        
    def find_producer_in_text(self, text):
        """Find the producer in the text using our known list"""
        # First check for exact matches
        for producer in self.producers:
            if producer.upper() in text:
                return producer
                
        # Then check for known variants
        for variant, full_name in self.producer_variants.items():
            if variant in text:
                return full_name
                
        return "UNKNOWN"
        
    def parse_description(self, description):
        """Parse the description into vintage, producer, and product components"""
        parts = description.split()
        
        # First part is usually the vintage (year or N/V)
        vintage = parts[0]
        if vintage.startswith('20') or vintage == 'N/V':
            parts = parts[1:]  # Remove vintage from parts
        else:
            vintage = ''
            
        # Find the producer using our list
        full_desc = ' '.join(parts)
        producer = self.find_producer_in_text(full_desc.upper())
        
        # Remove producer from description to get product name
        product = full_desc
        if producer != "UNKNOWN":
            # Remove the producer name (and any variants) from the product description
            product = product.replace(producer, '')
            for variant in self.producer_variants.keys():
                product = product.replace(variant, '')
        product = ' '.join(product.split())  # Clean up extra spaces
        
        return vintage, producer, product
        
    def extract_pdf_data(self, pdf_path):
        """Extract data from PDF using direct text extraction"""
        print(f"\nExtracting data from {pdf_path}...")
        data = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    lines = text.split('\n')
                    
                    for line in lines:
                        # Skip header lines and empty lines
                        if not line.strip() or 'DESCRIPTION' in line or 'PAGE' in line or 'RUN' in line:
                            continue
                            
                        # Each line should start with an 'S' followed by alphanumeric characters
                        if line[0] == 'S':
                            try:
                                # Split by spaces, but keep the description together
                                parts = line.split()
                                if len(parts) >= 4:  # Make sure we have at least SKU, description, and some numbers
                                    sku = parts[0]
                                    
                                    # The last three numbers are on_hand, on_order, available
                                    available = parts[-1]
                                    on_order = parts[-2]
                                    on_hand = parts[-3]
                                    
                                    # Everything in between is the description
                                    description = ' '.join(parts[1:-3])
                                    
                                    # Parse the description into components
                                    vintage, producer, product = self.parse_description(description)
                                    
                                    data.append({
                                        'SKU': sku,
                                        'Vintage': vintage,
                                        'Producer': producer,
                                        'Product': product,
                                        'Full Description': description,
                                        'On Hand': on_hand,
                                        'On Order': on_order,
                                        'Available': available,
                                        'On Website': False,
                                        'Has Bottle Shot': False,
                                        'Has Label': False,
                                        'Product URL': ''
                                    })
                            except Exception as e:
                                print(f"Skipping line due to error: {str(e)}")
                                continue
                                
        except Exception as e:
            print(f"Error extracting PDF data: {str(e)}")
            return None
            
        if not data:
            print("\nNo data was extracted from the PDF!")
            return None
            
        print(f"\nExtraction complete. Found {len(data)} products.")
        return pd.DataFrame(data)

    def setup_selenium(self):
        """Initialize Selenium WebDriver"""
        print("\nSetting up web browser...")
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.set_page_load_timeout(30)

    def wait_for_element(self, selector, timeout=10):
        """Wait for an element to be present on the page"""
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, selector))
            )
            return True
        except TimeoutException:
            return False
            
    def get_page_content(self):
        """Get all text content from the page for debugging"""
        return self.driver.execute_script("""
            return document.body.innerText;
        """)
        
    def get_producer_page_url(self, producer):
        """Convert producer name to URL format"""
        url_mapping = {
            'Mount Fishtail': 'mount-fishtail-wines',
            'Sherwood': 'sherwood-wines',
            'Stratum': 'stratum-wines',
            'Black Pearl Vineyards': 'black-pearl-wines',
            'Downes Family Vineyards': 'downes-family-vineyards-wines',
            'David Finlayson': 'david-finlayson-wines',
            'Luddite': 'luddite-wines',
            'Painted Wolf': 'painted-wolf-wines',
            'Thistledown': 'thistledown-wines',
            'Nugan Estate': 'nugan-wines',
            'RL Buller': 'rl-buller-wines',
            'Water Wheel': 'water-wheel-wines',
            'Casas del Bosque': 'casas-del-bosque-wines',
            'Earthsong': 'earthsong-wines',
            'Miha': 'miha-wines-2',
            'Almarada': 'almarada-wines',
            'Antucura': 'antucura-wines',
            'Vina Alicia': 'vina-alicia-2',
            'Greenock Creek': 'greenock-wines',
            'Oliverhill': 'oliver-hill-wines',
            'Tait': 'tait-wines',
            'Wildberry Estate': 'wildberry-estate-wines',
            'NEWBLOOD': 'newblood-wines',
            'Rieslingfreak': 'rieslingfreak-wines'
        }
        
        if producer in url_mapping:
            return f"{self.base_url}/{url_mapping[producer]}/"
        return None

    def get_producer_products(self, producer):
        """Get all products for a producer from the website"""
        producer_url = self.get_producer_page_url(producer)
        if not producer_url:
            print(f"No URL mapping found for producer: {producer}")
            return []
            
        try:
            print(f"\nChecking {producer_url}...")
            self.driver.get(producer_url)
            time.sleep(5)  # Give page time to load
            
            print(f"Page title: {self.driver.title}")
            print(f"Current URL: {self.driver.current_url}")
            
            # Look for wine products using Elementor's structure
            products = []
            
            # Find all article elements that are wine posts
            wine_elements = self.driver.find_elements(By.CSS_SELECTOR, 'article[type-wines]')
            print(f"Found {len(wine_elements)} wine elements")
            
            for element in wine_elements:
                try:
                    # Get the wine title
                    title = element.find_element(By.CSS_SELECTOR, '.elementor-heading-title').text
                    if not title:
                        continue
                        
                    # Get the View Wine button link
                    view_btn = element.find_element(By.CSS_SELECTOR, '.elementor-button-wrapper a')
                    url = view_btn.get_attribute('href')
                    
                    print(f"Found wine: {title}")
                    products.append({
                        'title': title,
                        'url': url
                    })
                    
                except NoSuchElementException:
                    continue
                except Exception as e:
                    print(f"Error processing wine element: {e}")
                    continue
                    
            # If no products found with article elements, try looking for buttons directly
            if not products:
                print("Trying alternate approach with View Wine buttons...")
                buttons = self.driver.find_elements(By.CSS_SELECTOR, '.elementor-button-wrapper a')
                for button in buttons:
                    try:
                        # Get the button's parent container to find the title
                        container = button.find_element(By.XPATH, './ancestor::article')
                        title = container.find_element(By.CSS_SELECTOR, '.elementor-heading-title').text
                        url = button.get_attribute('href')
                        
                        if title and url:
                            print(f"Found wine via button: {title}")
                            products.append({
                                'title': title,
                                'url': url
                            })
                            
                    except Exception as e:
                        continue
                        
            print(f"\nFinished processing {producer}. Found {len(products)} wines.")
            return products
            
        except Exception as e:
            print(f"Error getting products for {producer}: {e}")
            return []

    def check_product_details(self, product_url):
        """Check if a product has bottle shot and label images"""
        try:
            self.driver.get(product_url)
            time.sleep(1)
            
            images = self.driver.find_elements(By.CSS_SELECTOR, '.product-gallery__image')
            
            has_bottle = False
            has_label = False
            
            for img in images:
                src = img.get_attribute('src').lower()
                if 'bottle' in src:
                    has_bottle = True
                if 'label' in src:
                    has_label = True
                    
            return has_bottle, has_label
            
        except Exception as e:
            print(f"Error checking product details at {product_url}: {e}")
            return False, False

    def process_inventory(self, pdf_path):
        """Main method to process inventory and check website"""
        try:
            # Extract data from the PDF
            df = self.extract_pdf_data(pdf_path)
            
            if df is None or df.empty:
                raise Exception("No data extracted from PDF")
                
            print(f"\nFound {len(df)} products in PDF")
                
            # Set up Selenium for website checking
            self.setup_selenium()
            
            # Track results for both inventory matches and website-only products
            inventory_results = []
            website_only = []
            
            # Process each producer
            unique_producers = df['Producer'].unique()
            for producer in unique_producers:
                if producer == "UNKNOWN":
                    continue
                    
                print(f"\nProcessing {producer}...")
                website_products = self.get_producer_products(producer)
                
                # Track which website products were matched
                matched_products = set()
                
                # Update DataFrame for matching products
                producer_rows = df[df['Producer'] == producer]
                for _, row in producer_rows.iterrows():
                    sku = row['SKU']
                    vintage = row['Vintage']
                    product = row['Product']
                    
                    result = {
                        'SKU': sku,
                        'Vintage': vintage,
                        'Producer': producer,
                        'Product': product,
                        'Full Description': row['Full Description'],
                        'On Hand': row['On Hand'],
                        'On Order': row['On Order'],
                        'Available': row['Available'],
                        'On Website': False,
                        'Has Bottle Shot': False,
                        'Has Label': False,
                        'Product URL': ''
                    }
                    
                    # Try to find matching product
                    matching_product = None
                    for product_info in website_products:
                        # Check if SKU matches or if both vintage and product name match
                        if (sku in product_info['title'] or 
                            (vintage in product_info['title'] and any(word in product_info['title'] for word in product.split()))):
                            matching_product = product_info
                            matched_products.add(product_info['title'])
                            break
                            
                    if matching_product:
                        print(f"Found match for {sku} on website")
                        # Check for bottle shot and label
                        has_bottle, has_label = self.check_product_details(matching_product['url'])
                        
                        result.update({
                            'On Website': True,
                            'Has Bottle Shot': has_bottle,
                            'Has Label': has_label,
                            'Product URL': matching_product['url']
                        })
                        
                    inventory_results.append(result)
                    
                # Add website-only products (those not matched to inventory)
                for product_info in website_products:
                    if product_info['title'] not in matched_products:
                        # Check for bottle shot and label
                        has_bottle, has_label = self.check_product_details(product_info['url'])
                        
                        website_only.append({
                            'Producer': producer,
                            'Product Name': product_info['title'],
                            'Has Bottle Shot': has_bottle,
                            'Has Label': has_label,
                            'Product URL': product_info['url']
                        })
            
            # Create DataFrames
            inventory_df = pd.DataFrame(inventory_results)
            website_only_df = pd.DataFrame(website_only)
            
            # Sort both DataFrames
            inventory_df = inventory_df.sort_values(['Producer', 'SKU'])
            website_only_df = website_only_df.sort_values(['Producer', 'Product Name'])
            
            # Save results to Excel with multiple sheets
            output_file = 'inventory_report.xlsx'
            i = 1
            while True:
                try:
                    if i > 1:
                        base, ext = os.path.splitext(output_file)
                        current_file = f"{base}_{i}{ext}"
                    else:
                        current_file = output_file
                        
                    with pd.ExcelWriter(current_file, engine='openpyxl') as writer:
                        inventory_df.to_excel(writer, sheet_name='Inventory Report', index=False)
                        website_only_df.to_excel(writer, sheet_name='Website Only', index=False)
                        
                    print(f"\nResults saved to {current_file}")
                    print(f"- Inventory Report: {len(inventory_df)} products")
                    print(f"- Website Only: {len(website_only_df)} products")
                    break
                except PermissionError:
                    i += 1
                    if i > 10:
                        raise Exception("Could not save file after 10 attempts")
                    continue
            
            return True
            
        except Exception as e:
            print(f"Error processing inventory: {e}")
            return False
            
        finally:
            if self.driver:
                self.driver.quit()

def main():
    parser = argparse.ArgumentParser(description='Process inventory PDF and check website')
    parser.add_argument('pdf_path', help='Path to the inventory PDF file')
    args = parser.parse_args()
    
    if not os.path.exists(args.pdf_path):
        print(f"Error: File not found: {args.pdf_path}")
        return
        
    checker = InventoryChecker()
    checker.process_inventory(args.pdf_path)

if __name__ == "__main__":
    main() 