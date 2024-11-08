
# import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException
from bs4 import BeautifulSoup
from lxml import html
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import pandas as pd
import re
import os
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

class GoogleReviewsScraper:
    # def __init__(self, chromedriver_path, input_file_path, output_file_path):
    def __init__(self, input_file_path, output_file_path):
        # self.chromedriver_path = chromedriver_path
        self.input_file_path = input_file_path
        self.output_file_path = output_file_path
        self.driver = None
        self.wait = None
        self.wait_short = None
        self.current_date = datetime.now()
        
        self.date_patterns = {
            r"(\d+)\s*years?\s*ago": "years",
            r"(\d+)\s*months?\s*ago": "months",
            r"(\d+)\s*weeks?\s*ago": "weeks",
            r"(\d+)\s*days?\s*ago": "days",
            r"a\s*year\s*ago": "1 year",
            r"a\s*month\s*ago": "1 month",
            r"a\s*week\s*ago": "1 week",
            r"a\s*day\s*ago": "1 day",
            r"today": "today"
        }

    def setup_driver(self):
        options = Options()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-logging')
        options.add_argument('--enable-automation')
        options.add_argument('--log-level=3')
        options.add_argument('--v=99')
        options.add_argument('--headless')
        options.binary_location = '/usr/bin/google-chrome'
        self.driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
        # self.driver = uc.Chrome(driver_executable_path=self.chromedriver_path, options=options)

        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, 5)
        self.wait_short = WebDriverWait(self.driver, 1)

    def parse_date(self, date_text):
        specific_date = None
        for pattern, unit in self.date_patterns.items():
            match = re.search(pattern, date_text)
            if match:
                if unit == "today":
                    specific_date = self.current_date
                else:
                    amount = 1 if unit in ["1 year", "1 month", "1 week", "1 day"] else int(match.group(1))
                    
                    if unit in ["1 year", "years"]:
                        specific_date = self.current_date - relativedelta(years=amount)
                    elif unit in ["1 month", "months"]:
                        specific_date = self.current_date - relativedelta(months=amount)
                    elif unit in ["1 week", "weeks"]:
                        specific_date = self.current_date - timedelta(weeks=amount)
                    elif unit in ["1 day", "days"]:
                        specific_date = self.current_date - timedelta(days=amount)
                break
        
        if specific_date:
            return specific_date.strftime("%m/%d/%Y"), specific_date.strftime("%Y")
        return 'N/A', 'N/A'

    def load_all_reviews(self, total_expected_reviews):
        previous_review_count = 0
        max_attempts = 50
        attempts = 0

        while attempts < max_attempts:
            try:
                body_texts = self.wait.until(EC.visibility_of_all_elements_located(
                    (By.XPATH, '//div[@jscontroller="fIQYlf"]')))
                current_review_count = len(body_texts)
                
                if current_review_count >= total_expected_reviews:
                    print(f"Successfully loaded all {current_review_count} reviews")
                    break
                    
                if current_review_count == previous_review_count:
                    attempts += 1
                else:
                    attempts = 0
                    previous_review_count = current_review_count
                
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                sleep(1)
                
                try:
                    scroll_element = self.wait.until(EC.visibility_of_element_located((By.XPATH, '//div[@class="loris"]')))
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", scroll_element)
                    self.wait.until(EC.invisibility_of_element_located((By.XPATH, '//div[@class="jfk-activityIndicator"]')))
                except (StaleElementReferenceException, NoSuchElementException):
                    continue
                
                print(f"Loaded {current_review_count}/{total_expected_reviews} reviews...")
                
            except Exception as e:
                print(f"Error loading reviews: {e}")
                continue

    def expand_review_texts(self):
        try:
            more_buttons = self.driver.find_elements(By.XPATH, '//a[@class="review-more-link"]')
            for button in more_buttons:
                self.driver.execute_script("arguments[0].scrollIntoView(true);", button)
                self.driver.execute_script("arguments[0].click();", button)
        except Exception as e:
            print(f"Error expanding reviews: {e}")

    def extract_review_data(self, location_data):
        reviews_data = []
        
        html_content = self.driver.page_source
        soup = BeautifulSoup(html_content, 'html.parser')
        card_elements = soup.select('div.WMbnJf.vY6njf.gws-localreviews__google-review')
        
        total_reviews_element = self.driver.find_elements(By.XPATH, '//span[@class="z5jxId"]')
        total_reviews = int(re.search(r'\d+', total_reviews_element[0].text.strip()).group()) if total_reviews_element else 0
        
        ratings_element = self.driver.find_elements(By.XPATH, '//g-review-stars/span[@aria-label]')
        rating_match = re.search(r'(\d\.\d) out of (\d)', ratings_element[0].get_attribute('aria-label')) if ratings_element else None
        rating_value = rating_match.group(1) if rating_match else 'N/A'
        rating_max = rating_match.group(2) if rating_match else 'N/A'

        for card_element in card_elements:
            element = html.fromstring(str(card_element))
            
            review_texts = element.xpath('//div[@class="Jtu6Td"]')
            for i in range(0, len(review_texts), 2):
                review_text = ' '.join(line.strip() for line in review_texts[i].text_content().splitlines() if line.strip())
                
                date_element = element.xpath('//span[@class="dehysf lTi8oc"]')[i].text_content().strip()
                review_date, review_year = self.parse_date(date_element)
                
                reviews_data.append({
                    **location_data,
                    'Rating left by reviewer (out of 5)': rating_max,
                    'Review date': review_date,
                    'Average Review for location': rating_value,
                    'Year of review': review_year,
                    '# of Reviews for location': total_reviews,
                    'Review text (local language)': review_text
                })
        
        return reviews_data

    def scrape_reviews(self):
        try:
            df_input = pd.read_excel(self.input_file_path)
            
            country_, region_, reviewDate_, name_, location_ = [], [], [], [], []
            cohort_, out_of_value, comments_time, rating_value = [], [], [], []
            years, total_reviews, address_, reviews_texts = [], [], [], []
            googleId_, googleLinks_, ones_, metro_ = [], [], [], []

            for _, row in df_input.iterrows():
                self.driver.get(row['Link'])

                try:
                    not_now = self.wait.until(EC.visibility_of_element_located((By.XPATH, '//*[contains(text(),"Not now")]')))
                    self.driver.execute_script("arguments[0].click();", not_now)
                except:
                    pass
                try:
                    newest = self.wait_short.until(EC.visibility_of_element_located((By.XPATH, '//*[contains(text(), "Newest")]')))
                    self.driver.execute_script("arguments[0].click();", newest)
                except Exception as e:
                    print(f"Could not switch to newest reviews: {e}")

                reviews_count = self.driver.find_elements(By.XPATH, '//span[@class="z5jxId"]')
                if reviews_count:
                    total_expected = int(re.search(r'\d+', reviews_count[0].text.strip()).group())
                    self.load_all_reviews(total_expected)
                    self.expand_review_texts()

                    html_content = self.driver.page_source
                    soup = BeautifulSoup(html_content, 'html.parser')
                    card_elements = soup.select('div.WMbnJf.vY6njf.gws-localreviews__google-review')

                    ratings_element = self.driver.find_elements(By.XPATH, '//g-review-stars/span[@aria-label]')
                    rating_match = re.search(r'(\d\.\d) out of (\d)', ratings_element[0].get_attribute('aria-label')) if ratings_element else None
                    rating = rating_match.group(1) if rating_match else 'N/A'
                    max_rating = rating_match.group(2) if rating_match else 'N/A'

                    for card_element in card_elements:
                        element = html.fromstring(str(card_element))
                        
                        review_texts = element.xpath('//div[@class="Jtu6Td"]')
                        date_elements = element.xpath('//span[@class="dehysf lTi8oc"]')
                        
                        for i in range(0, len(review_texts), 2):
                            review_text = ' '.join(line.strip() for line in review_texts[i].text_content().splitlines() if line.strip())
                            date_text = date_elements[i].text_content().strip()
                            review_date, review_year = self.parse_date(date_text)

                            country_.append(row['Country'])
                            region_.append(row['Province / region'])
                            reviewDate_.append(row['Review date'])
                            name_.append(row['Name of chain'])
                            location_.append(row['Name Of Location'])
                            cohort_.append(row['Cohort year'])
                            out_of_value.append(max_rating)
                            comments_time.append(review_date)
                            rating_value.append(rating)
                            years.append(review_year)
                            total_reviews.append(total_expected)
                            address_.append(row['Address'])
                            reviews_texts.append(review_text)
                            googleId_.append(row['Google ID'])
                            googleLinks_.append(row['Link'])
                            ones_.append(row['One'])
                            metro_.append(row['Metro Area'])

            data = zip(country_, region_, reviewDate_, name_, location_, cohort_,
                    out_of_value, comments_time, rating_value, years, total_reviews,
                    address_, reviews_texts, googleId_, googleLinks_, ones_, metro_)
            
            headings = ['Country', 'Province / region', 'Review date', 'Name of chain',
                    'Name Of Locatione', 'Cohort year', 'Rating left by reviewer (out of 5)',
                    'Review date', 'Average Review for location', 'Year of review',
                    '# of Reviews for location', 'Address', 'Review text (local language)',
                    'Google ID', 'Link', '1', 'Metro Area']

            df = pd.DataFrame(data, columns=headings)
            df.to_excel(self.output_file_path, index=False)

        except Exception as e:
            print(f"Error during scraping: {e}")
        finally:
            if self.driver:
                self.driver.quit()
    

def main():
    # chromedriver_path = r"E:\\chromedriver.exe"
    input_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'review.xlsx')
    output_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Google Review.xlsx')
    # output_file_path = r"C:\Users\Mauz Khan\Desktop\Files\Google Review\Google Review.xlsx"
    
    scraper = GoogleReviewsScraper(input_file_path, output_file_path)
    # scraper = GoogleReviewsScraper(chromedriver_path, input_file_path, output_file_path)
    scraper.setup_driver()
    scraper.scrape_reviews()

if __name__ == "__main__":
    main()