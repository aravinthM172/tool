import time
from datetime import datetime,timedelta
import requests
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from amazoncaptcha import AmazonCaptcha

currenttime = datetime.strftime(datetime.utcnow(), '%Y-%m-%dT%H:%M:%S.000Z')
# Set up the Selenium WebDriver
service = Service(executable_path='/usr/bin/chromedriver')
user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36'
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
chrome_options.add_argument(f'user-agent={user_agent}')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--ignore-certificate-errors-spki-list')
chrome_options.add_argument('--ignore-ssl-errors')
chrome_options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(options=chrome_options)

# Define a function to obtain Amazon cookies
def amazon_cookies(test):
    repeat = True
    attempt = 0
    amazon_cookie = ''
    while repeat == True:
        urlamazonin = 'https://www.amazon.in/'
        headers = {
            'client-type': 'd-web',
            'User-Agent': 'PostmanRuntime/7.29.0'
        }
        response = requests.request("GET", urlamazonin, headers=headers)

        cookies = response.cookies
        data = []
        for cookie in cookies:
            data.append(cookie.name + '=' + cookie.value + ';')
        amazon_cookie = ''.join([str(x) for x in data])
        print('amazon_cookies: ', amazon_cookie)
        if amazon_cookie:
            repeat = False
        else:
            attempt += 1
            if attempt <= 10:
                time.sleep(2)
                if attempt >= 3:
                    time.sleep(30)
                repeat = True
            else:
                print('max_attempts')
                repeat = False
    return amazon_cookie
all_product_data_list = []
# Define a function to fetch global rating and reviews for a product
def global_rating_reviews(product_id, amazon_cookie):
    attempt = 0
    attempt1 = 0
    product_data_list = []
    try:
        url = 'https://www.amazon.in/product-reviews/{}?ie=UTF8&reviewerType=all_reviews&formatType=current_format&&sortBy=recent'.format(
            product_id)
        repeat = True
        global_reviews = '0'
        global_ratings = '0'
        global_stars = '0'
        attempt = 0
        attempt1 = 0
        while repeat == True:
            driver.get(url)
            response = driver.page_source
            status_code = 200
            print('status_code: ', status_code)
            if 200 == status_code:
                soup = BeautifulSoup(response, 'html.parser')
                all_stars = soup.find(attrs={'data-hook': 'rating-out-of-text'}).get_text() if soup.find(
                    attrs={'data-hook': 'rating-out-of-text'}) else 0
                global_rating_reviews = soup.find(
                    attrs={'data-hook': 'cr-filter-info-review-rating-count'}).get_text().strip() if soup.find(
                    attrs={'data-hook': 'cr-filter-info-review-rating-count'}) else ''
                print(all_stars, global_rating_reviews)
                if all_stars and global_rating_reviews:
                    global_stars = all_stars.split(' ')[0].lstrip().replace(',', '')
                    global_rating_reviews = global_rating_reviews.split('ratings,')
                    global_ratings = global_rating_reviews[0].split(' ')[0].lstrip().replace(',', '') if len(
                        global_rating_reviews) > 0 else 0
                    global_reviews = global_rating_reviews[1].lstrip().split(' ')[0].lstrip().replace(',', '') if len(
                        global_rating_reviews) > 1 else 0
                    # print({"Product_id": product_id, "Reviews": global_reviews, "Rating": global_ratings,
                    #        "Stars": global_stars})

                    product_data = {
                        "ASIN": product_id,
                        "Global Reviews": global_reviews,
                        "Global Rating": global_ratings,
                        "Stars": global_stars,
                        "Date an" : currenttime
                    }

                    print("product_data:", product_data)

                    product_data_list.append(product_data)
                    product_df = pd.DataFrame(product_data_list)

                    return product_df
                    repeat = False
                else:
                    driver.get('https://www.amazon.in/errors/validateCaptcha')
                    captcha = AmazonCaptcha.fromdriver(driver)
                    solution = captcha.solve()
                    attempt1 += 1
                    if attempt1 <= 10:
                        time.sleep(2)
                        if attempt1 >= 5:
                            time.sleep(30)
                            print(url)
                            print('URL_getting= {} error, Attempt_count = {}, Product_id= {},org_id = {}'.format(
                                status_code, attempt1, product_id, org_id))
                        repeat = True
                    else:
                        print('max_attempts')
                        repeat = False

            elif 404 == status_code:
                print('URL_getting= {} error,org_id = {}, Product_id = {}'.format(status_code, org_id, product_id))
                repeat = False
            else:
                attempt += 1
                if attempt <= 10:
                    time.sleep(2)
                    if attempt >= 3:
                        time.sleep(30)
                        print(url)
                        print(
                            'URL_getting= {} error, Attempt_count = {}, Product_id= {},org_id = {}'.format(status_code,
                                                                                                           attempt,
                                                                                                           product_id,
                                                                                                           org_id))
                    repeat = True
                else:
                    print('max_attempts')
                    repeat = False

    except Exception as ex:
        print('GlobalReviewRatings Error')
        amz_cookie = amazon_cookies(test='test')
        repeat = False
        print('error_details: ', ex)
        attempt += 1
        if attempt <= 5:
            print('hello')
            time.sleep(2)
            if attempt >= 3:
                time.sleep(30)
                print(url)
                print(
                    'URL_getting= {} error, Attempt_count = {}, Product_id= {},org_id = {}'.format(status_code, attempt,
                                                                                                   product_id, org_id))
            repeat = True
        else:
            print('max_attempts')
            repeat = False

# Define the input and output file paths
amz_cookie = "test"
input_excel_file = "sampleasin.xlsx"
output_excel_file = "output_product_info2.xlsx"

# Read the product IDs from the input Excel file
df = pd.read_excel(input_excel_file)
product_ids = df['ASIN']

# Create a list to store the output data
product_info_list = []

# Iterate through the product IDs and fetch their global rating and reviews
for product_id in product_ids:
    product_data = global_rating_reviews(product_id=product_id, amazon_cookie=amz_cookie)
    if product_data is not None:
        all_product_data_list.append(product_data)

final_product_df = pd.concat(all_product_data_list, ignore_index=True)
print(final_product_df)
# Create a DataFrame from the product_info_list
# output_df = pd.DataFrame(product_info_list, columns=["Product_id", "Reviews", "Rating", "Stars"])
#
# # Save the DataFrame to an Excel file
final_product_df.to_excel(output_excel_file, index=False)
