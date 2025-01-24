import os
import logging
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from bs4 import BeautifulSoup
import pandas as pd

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def scroll_smoothly(driver, element, step=100, delay=0.1, timeout=40):
    """
    Smoothly scrolls the given element by a specified step size with a delay between each step.
    Stops scrolling if the scroll height does not change for a specified timeout period.

    :param driver: Selenium WebDriver instance
    :param element: WebElement to scroll
    :param step: Number of pixels to scroll in each step
    :param delay: Delay in seconds between each scroll step
    :param timeout: Timeout in seconds to wait for scroll height to change
    """
    try:
        # Get the total scroll height of the element
        scroll_height = driver.execute_script("return arguments[0].scrollHeight", element)
        # Get the current scroll position
        current_position = driver.execute_script("return arguments[0].scrollTop", element)

        last_scroll_height = 0
        start_time = time.time()

        while True:
            # Scroll by the specified step
            driver.execute_script(f"arguments[0].scrollTop += {step};", element)
            # Update the current position
            current_position = driver.execute_script("return arguments[0].scrollTop", element)
            # Log the current scroll position
            logging.info(f"Scrolled to position: {current_position}/{scroll_height}")
            # Wait for the specified delay
            time.sleep(delay)

            # Check if the scroll height has changed
            new_scroll_height = driver.execute_script("return arguments[0].scrollHeight", element)
            if new_scroll_height != last_scroll_height:
                last_scroll_height = new_scroll_height
                start_time = time.time()  # Reset the timer if scroll height changes
            else:
                # Check if the timeout period has been reached
                if time.time() - start_time > timeout:
                    logging.info("Scroll height did not change for the specified timeout period. Stopping scrolling.")
                    break

            # Update the scroll height
            scroll_height = new_scroll_height

        logging.info("Reached the end of the scrollable container.")

    except Exception as e:
        logging.error(f"An error occurred while scrolling smoothly: {e}")



def get_reviews_from_YandexMaps(bs4_soup):
    reviews = []
    authors_ratings = []
    publication_date = []

    for result in bs4_soup.findAll('div', class_='business-reviews-card-view__review'):
        try:
            reviews.append(result.find('span', class_='business-review-view__body-text').get_text())
        except AttributeError:
            reviews.append(pd.NA)

        try:
            rating_value = result.find('div', class_='business-review-view__rating').find('span', attrs={'itemprop': 'reviewRating'}).find('meta', attrs={'itemprop': 'ratingValue'})
            authors_ratings.append(float(rating_value['content']) if rating_value else pd.NA)
        except (AttributeError, TypeError):
            authors_ratings.append(pd.NA)

        try:
            date_meta = result.find('span', class_='business-review-view__date').find('meta')
            publication_date.append(date_meta['content'] if date_meta else None)
        except (AttributeError, TypeError):
            publication_date.append(None)

    data = pd.DataFrame.from_dict({
        'rating': authors_ratings,
        'publication_date': publication_date,
        'review': reviews
    })

    data['publication_date'] = pd.to_datetime(data['publication_date'], errors='coerce').dt.date

    return data

def main(url):
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")

    # Use environment variable for chromedriver path
    chromedriver_path = os.getenv('CHROMEDRIVER_PATH', 'C:/Program Files/Google/Chrome/Application/chromedriver-win64/chromedriver.exe')
    service = Service(chromedriver_path)
    browser = webdriver.Chrome(service=service, options=chrome_options)

    try:
        browser.get(url)

        # Ensure page is fully loaded before proceeding
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".scroll__container")))

        # Find the scrollable container
        site_body = WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".scroll__container")))

        # Scroll smoothly to the bottom
        scroll_smoothly(driver=browser, element=site_body)

        content = browser.page_source
        soup = BeautifulSoup(content, features='html.parser')
        data = get_reviews_from_YandexMaps(bs4_soup=soup)

        output_file = f"reviews.xlsx"
        output_path = os.path.join("output", output_file)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        data.to_excel(output_path, index=False)

        logging.info('Work is DONE!')

    except Exception as e:
        logging.error(f"An error occurred while processing {url}: {e}")
    finally:
        browser.quit()

if __name__ == '__main__':
    # Directly specify the URL to scrape
    url = ''
    main(url)
