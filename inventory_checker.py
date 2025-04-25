import os
import pandas as pd
import pdfplumber
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
import argparse

class InventoryChecker:
    def __init__(self):
        self.base_url = "https://southernstarz.com"
        self.driver = None
        self.max_retries = 3
        self.page_load_timeout = 30
        self.wait = None
        
        # Add direct SKU to URL mappings
        self.sku_url_mapping = {
            'S1ANTCHNV': 'https://southernstarz.com/wines/antucura-cherie-sparkling-nv/',
            'S1NBCHANAC': 'https://southernstarz.com/wines/newblood-chardonnay/',
            'S1NBCHANAE': 'https://southernstarz.com/wines/newblood-chardonnay/',
            'S1NBCHANAF': 'https://southernstarz.com/wines/newblood-chardonnay/',
            'S1NBRBLNAE': 'https://southernstarz.com/wines/newblood-red-blend/',
            'S1NBRBLNAF': 'https://southernstarz.com/wines/newblood-red-blend/',
            'S1NBROSNAF': 'https://southernstarz.com/wines/newblood-rose/'
        }
        
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
            'GREENOCK': 'Greenock Creek',
            'NEW BLOOD': 'NEWBLOOD',
            'NEWBLOOD NON-ALCOHOLIC': 'NEWBLOOD',
            'ANTUCURA CHERIE': 'Antucura'
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
                                        'Has Spec Sheet': False,
                                        'Has Shelf-Talker': False,
                                        'Has Hi-Res Label': False,
                                        'Has Bottle Shot': False,
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
        """Initialize Selenium WebDriver with retry logic"""
        print("\nSetting up web browser...")
        for attempt in range(self.max_retries):
            try:
                if self.driver is not None:
                    try:
                        self.driver.quit()
                    except:
                        pass
                options = webdriver.ChromeOptions()
                options.add_argument('--headless')
                options.add_argument('--disable-gpu')
                options.add_argument('--no-sandbox')
                options.add_argument('--disable-dev-shm-usage')
                options.add_argument('--window-size=1920,1080')
                # Performance optimizations
                options.add_argument('--disable-extensions')
                options.add_argument('--disable-images')
                options.add_argument('--disable-javascript')
                options.add_argument('--blink-settings=imagesEnabled=false')
                options.page_load_strategy = 'eager'
                service = Service(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service, options=options)
                self.driver.set_page_load_timeout(self.page_load_timeout)
                self.wait = WebDriverWait(self.driver, 10)
                return True
            except Exception as e:
                print(f"Attempt {attempt + 1} failed: {str(e)}")
                time.sleep(2)
        return False

    def safe_get_url(self, url, max_retries=None):
        """Safely navigate to a URL with retry logic"""
        if max_retries is None:
            max_retries = self.max_retries
            
        for attempt in range(max_retries):
            try:
                if not self.driver or not self.is_session_valid():
                    if not self.setup_selenium():
                        return False
                print(f"\nLoading {url}...")
                self.driver.get(url)
                self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                return True
            except Exception as e:
                print(f"Attempt {attempt + 1} failed to load {url}: {str(e)}")
                time.sleep(2)
                if "timeout" in str(e).lower():
                    print("Page load timed out, retrying with longer timeout...")
                    try:
                        self.driver.set_page_load_timeout(self.page_load_timeout * 1.5)
                    except:
                        pass
        return False

    def is_session_valid(self):
        """Check if the current session is valid"""
        try:
            # Try to access a simple property
            _ = self.driver.current_url
            return True
        except:
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
        """Get all products for a producer from the website with improved error handling"""
        producer_url = self.get_producer_page_url(producer)
        if not producer_url:
            print(f"No URL mapping found for producer: {producer}")
            return []
            
        try:
            print(f"\nChecking {producer_url}...")
            if not self.safe_get_url(producer_url):
                return []
                
            print(f"Page title: {self.driver.title}")
            print(f"Current URL: {self.driver.current_url}")
            
            products = []
            
            # Wait for either wine elements or buttons to be present
            try:
                elements = self.wait.until(
                    EC.presence_of_all_elements_located((
                        By.CSS_SELECTOR, 
                        'article[type-wines], .elementor-button-wrapper a'
                    ))
                )
                
                # Process wine elements
                wine_elements = self.driver.find_elements(By.CSS_SELECTOR, 'article[type-wines]')
                if wine_elements:
                    for element in wine_elements:
                        try:
                            title = element.find_element(By.CSS_SELECTOR, '.elementor-heading-title').text
                            if not title:
                                continue
                                
                            view_btn = element.find_element(By.CSS_SELECTOR, '.elementor-button-wrapper a')
                            url = view_btn.get_attribute('href')
                            
                            print(f"Found wine: {title}")
                            products.append({
                                'title': title,
                                'url': url
                            })
                            
                        except Exception as e:
                            continue
                
                # If no products found, try buttons approach
                if not products:
                    print("Trying alternate approach with View Wine buttons...")
                    buttons = self.driver.find_elements(By.CSS_SELECTOR, '.elementor-button-wrapper a')
                    for button in buttons:
                        try:
                            container = button.find_element(By.XPATH, './ancestor::article')
                            title = container.find_element(By.CSS_SELECTOR, '.elementor-heading-title').text
                            url = button.get_attribute('href')
                            
                            if title and url:
                                print(f"Found wine via button: {title}")
                                products.append({
                                    'title': title,
                                    'url': url
                                })
                                
                        except Exception:
                            continue
                            
            except TimeoutException:
                print(f"Timeout waiting for products on {producer_url}")
                
            print(f"\nFinished processing {producer}. Found {len(products)} wines.")
            return products
            
        except Exception as e:
            print(f"Error getting products for {producer}: {e}")
            return []

    def check_product_details(self, product_url):
        """Check if a product has all required assets with retry logic"""
        results = {
            'Has Spec Sheet': False,
            'Has Shelf-Talker': False,
            'Has Hi-Res Label': False,
            'Has Bottle Shot': False
        }

        if not self.safe_get_url(product_url):
            return results

        try:
            # Wait for trade tools to be present
            try:
                trade_tools = self.wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.elementor-button-wrapper a'))
                )
            except TimeoutException:
                return results

            # Instead of clicking, we'll check the href attributes directly
            for tool in trade_tools:
                try:
                    tool_text = tool.text.strip().lower()
                    href = tool.get_attribute('href')
                    
                    if not href:
                        continue
                        
                    tool_map = {
                        'spec sheet': 'Has Spec Sheet',
                        'shelf-talker': 'Has Shelf-Talker',
                        'hi-res label': 'Has Hi-Res Label',
                        'bottle shot': 'Has Bottle Shot'
                    }

                    for text, key in tool_map.items():
                        if text in tool_text:
                            # If we find a matching button with an href, consider it available
                            results[key] = True
                            break

                except Exception:
                    continue

            return results

        except Exception as e:
            print(f"Error checking product details at {product_url}: {e}")
            return results

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
                        'Has Spec Sheet': False,
                        'Has Shelf-Talker': False,
                        'Has Hi-Res Label': False,
                        'Has Bottle Shot': False,
                        'Product URL': ''
                    }
                    
                    # First check if we have a direct SKU to URL mapping
                    if sku in self.sku_url_mapping:
                        print(f"Found direct URL mapping for {sku}")
                        result['On Website'] = True
                        result['Product URL'] = self.sku_url_mapping[sku]
                        # Check for bottle shot and label
                        asset_results = self.check_product_details(self.sku_url_mapping[sku])
                        result.update(asset_results)
                    else:
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
                            results = self.check_product_details(matching_product['url'])
                            
                            result.update({
                                'On Website': True,
                                'Has Spec Sheet': results['Has Spec Sheet'],
                                'Has Shelf-Talker': results['Has Shelf-Talker'],
                                'Has Hi-Res Label': results['Has Hi-Res Label'],
                                'Has Bottle Shot': results['Has Bottle Shot'],
                                'Product URL': matching_product['url']
                            })
                        
                    inventory_results.append(result)
                    
                # Add website-only products (those not matched to inventory)
                for product_info in website_products:
                    if product_info['title'] not in matched_products:
                        # Check for bottle shot and label
                        results = self.check_product_details(product_info['url'])
                        
                        website_only.append({
                            'Producer': producer,
                            'Product Name': product_info['title'],
                            'Has Spec Sheet': results['Has Spec Sheet'],
                            'Has Shelf-Talker': results['Has Shelf-Talker'],
                            'Has Hi-Res Label': results['Has Hi-Res Label'],
                            'Has Bottle Shot': results['Has Bottle Shot'],
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