import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import os

# Update the path to your msedgedriver.exe
service = EdgeService(executable_path='C:/Users/robot/Downloads/Programs/zillowscrap/Driver/msedgedriver.exe')

# Setup Edge options
options = Options()
options.add_argument(
    'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (HTML, like Gecko) Chrome/91.0.4472.124 '
    'Safari/537.36')
options.add_argument('--disable-blink-features=AutomationControlled')  # Hide WebDriver flag

# Initialize WebDriver
driver = webdriver.Edge(service=service, options=options)

# List to store product details
product_list = []

# Define file paths
temp_file_path = 'zappos_products_temp.xlsx'
final_file_path = 'zappos_products.xlsx'

try:
    # Open the Zappos website
    driver.get('https://www.zappos.com/womens-shoes')

    base_url = 'https://www.zappos.com'

    while True:
        # Wait for the page to load completely
        time.sleep(5)  # Adjust as needed based on page load times

        # Get the page source
        page_source = driver.page_source

        # Parse the page source with BeautifulSoup
        soup = BeautifulSoup(page_source, 'html.parser')

        # Find all <article> elements with class "_5-z Yx-z" and required attributes
        articles = soup.find_all('article', class_='_5-z Yx-z', attrs={'data-style-id': True})

        # Extract detailed information from each article
        for article in articles:
            # Find the <a> tag within the article
            a_tag = article.find('a', class_='zn-z')

            if a_tag:
                # Extract details from <dl> within the <a> tag
                dl = a_tag.find('dl')
                if dl:
                    # Initialize a dictionary to store the details
                    product_details = {}

                    # Extract and store details
                    for dt, dd in zip(dl.find_all('dt'), dl.find_all('dd')):
                        key = dt.get_text(strip=True)
                        value = dd.get_text(strip=True)
                        product_details[key] = value

                    # Extract the relative URL and prepend the base URL
                    relative_url = a_tag.get('href')
                    product_details['URL'] = base_url + relative_url if relative_url else None

                    # Extract the image from the <figure> tag
                    figure = article.find('figure')
                    if figure:
                        # Find <meta> tag within <figure>
                        meta_tag = figure.find('meta', itemprop='image')
                        if meta_tag:
                            product_details['Image URL'] = meta_tag.get('content')

                    # Append product details to the list
                    product_list.append(product_details)

        # Periodically save the data
        df = pd.DataFrame(product_list)
        df.to_excel(temp_file_path, index=False, engine='openpyxl')
        print("Temporary data saved to zappos_products_temp.xlsx")

        # Check for the "Next Page" button
        try:
            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.Wk-z[rel="next"]'))
            )
            next_button.click()
        except Exception as e:
            print("No more pages or error encountered:", e)
            break  # Exit loop if no "Next" button found or an error occurs

except KeyboardInterrupt:
    print("Script interrupted. Saving the collected data...")
    # Save data before exiting
    df = pd.DataFrame(product_list)
    df.to_excel(temp_file_path, index=False, engine='openpyxl')
    print("Temporary data saved to zappos_products_temp.xlsx")

finally:
    # Quit the browser
    driver.quit()
    # Rename the temporary file to the final Excel file
    if os.path.exists(temp_file_path):
        os.rename(temp_file_path, final_file_path)
        print("Final data has been saved to zappos_products.xlsx")
    else:
        print("No data to save.")
