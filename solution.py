import argparse

import pandas as pd
import os.path
import time
import csv


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

role_filters = ["Directeur Du système d'information", "DSI", "Head Of Information Technology",
"Directeur Des Services Numériques",
"Group CIO",
"IT Maanager",
"Directeur Systèmes D'Information",
"CISO",
"Informatique",
"IT",
"RSSI",
"Security",
"Sécurité"]
connexion_url = 'https://app.cognism.com/auth/sign-in'
target_file = 'Cogmism.xlsx'
input_sheet = 'Input'
email = 'anis.chebbi@cloudshift.fr'
pwd = 'Shift2Cloud@24'
selectors = {'email': "#single-spa-application\:\@cognism\/app-auth > app-auth-root > app-auth > div > app-auth-login > div > div > div > div:nth-child(2) > form > nz-form-item:nth-child(2) > nz-form-control > div > div > input",
             'teamspassword': 'i0118',
             'tngovpassword': 'passwordInput',
             'next': 'idSIButton9',
             'signin': 'idSIButton9',
             'search': '#people-picker-input',
             'convbutton': 'app-bar-86fcd49b-61a2-4701-b771-54728cd291fb'}
import selenium


edge_driver_path = os.path.join(os.getcwd(), 'msedgedriver.exe')


def read_file(file, sheet_name, country='France'):
    df = pd.read_excel(file, sheet_name=sheet_name)
    old_results_df = pd.read_csv('done.csv', delimiter=';')

    # Filter the DataFrame based on the provided country
    filtered_df = df[df['country'].str.lower() == country.lower()]

    # Get the list of account names
    account_names = filtered_df['Account Name'].tolist()

    # Filter out account names present in old_results_df
    account_names = [name for name in account_names if name not in old_results_df['Company'].tolist()]

    return account_names


class UserAgent:

    def __init__(self, email, password, companies, country, chunk_size=10, exact_match=False):
        self.driver = None
        self.exact_match = exact_match
        self.setup_driver()
        self.email = email
        self.companies = [companies[i:i+chunk_size] for i in range(0, len(companies), chunk_size)]
        self.country = country
        self.password = password
        self.companies_filter_open = False
        self._connect(email, password)

    def setup_driver(self):
        driver_options = Options()
        driver_options.add_argument("--disable-extensions")
        driver_options.add_argument("--disable-popup-blocking")
        driver_options.add_argument("--disable-infobars")
        driver_options.add_argument("--disable-blink-features=AutomationControlled")
        driver_options.add_argument("--disable-save-password-bubble")
        driver_options.add_argument("--start-maximized")
        driver_options.add_argument("--inprivate")
        edge_service = Service(edge_driver_path)
        self.driver: webdriver.Edge = webdriver.Edge(service=edge_service, options=driver_options)

    def quit(self):
        self.driver.quit()
        time.sleep(2)

    def fetch_element(self, by, selector, wait_time=60):
        return WebDriverWait(self.driver, wait_time).until(
            EC.element_to_be_clickable((by, selector))
        )

    def _connect(self, email, password):

        self.driver.get(connexion_url)

        # Wait for email input field to be present
        email_input = WebDriverWait(self.driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, selectors['email']))
        )
        email_input.send_keys(email)
        time.sleep(2)
        password_input = WebDriverWait(self.driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#single-spa-application\:\@cognism\/app-auth > app-auth-root > app-auth > div > app-auth-login > div > div > div > div:nth-child(2) > form > nz-form-item:nth-child(3) > nz-form-control > div > div > nz-input-group > input'))
        )

        # Enter the password
        password_input.send_keys(password)
        time.sleep(2)
        # Click on the submit button
        submit_button_id = '#single-spa-application\:\@cognism\/app-auth > app-auth-root > app-auth > div > app-auth-login > div > div > div > div:nth-child(2) > form > div.t-mt-5 > button'
        submit_button = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, submit_button_id))
        )
        submit_button.click()
        WebDriverWait(self.driver, 60).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#nav-container > section > div.t-h-full.t-bg-light-100.t-flex.t-items-center.t-justify-start.ng-tns-c73-0 > a._menu-item.ng-tns-c73-0._active-menu-item')))
        time.sleep(0.2)
        self.driver.get('https://app.cognism.com/search/prospects/company-search')

    def set_countries(self):
        self.fetch_element(By.CSS_SELECTOR,
                           '#app-container > app-prospects > section > div > app-search-company-search > div > div:nth-child(1) > app-search-company-search-filter > div.search-filter-wrap.t-text-primary-850.t-pb-5 > div.collapse-header.pendo-DoE7v8PW6Llkqnj210qr.ng-star-inserted').click()
        countries_input = self.fetch_element(By.CSS_SELECTOR,
                                             '#app-container > app-prospects > section > div > app-search-company-search > div > div:nth-child(1) > app-search-company-search-filter > div.search-filter-wrap.t-text-primary-850.t-pb-5 > div.t-bg-light-200.ng-star-inserted > app-search-company-multiple-location-filters > nz-tabset > div > div > div.ant-tabs-tabpane.ant-tabs-tabpane-active.ng-star-inserted > app-search-company-multiple-all-location-filters > div > div:nth-child(2) > app-search-cognism-ng-select > div > ng-select > div > div > div.ng-input > input[type=text]')

        countries_input.click()
        countries_input.send_keys(self.country)
        time.sleep(1)
        #countries_input.send_keys(Keys.ARROW_DOWN)
        #time.sleep(0.2)
        countries_input.send_keys(Keys.ENTER)
        time.sleep(5)

    def toggle_exactmatch(self):
        if not self.exact_match:
            return
        # Toggle exact match
        WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                                         '#app-container > app-prospects > section > div > app-search-company-search > div > div:nth-child(1) > app-search-company-search-filter > div.search-filter-wrap.t-text-primary-850.t-pb-5 > div.collapse-header.pendo-4bH8v6F7o1Ac083UlIWj'))).click()

        WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.CSS_SELECTOR,

                                                                         '#app-container > app-prospects > section > div > app-search-company-search > div > div:nth-child(1) > app-search-company-search-filter > div.search-filter-wrap.t-text-primary-850.t-pb-5 > div.collapse-content.t-text-xs.pendo-7FhOg7f5Xge22BmA36tu.ng-star-inserted > div:nth-child(1) > nz-switch > button'))).click()

    def set_companies(self, chunk):
        if not self.companies_filter_open:
            self.companies_filter_open = True
            WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                                             '#app-container > app-prospects > section > div > app-search-company-search > div > div:nth-child(1) > app-search-company-search-filter > div.search-filter-wrap.t-text-primary-850.t-pb-5 > div:nth-child(2)'))).click()

        company_names_input = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                                                               '#app-container > app-prospects > section > div > app-search-company-search > div > div:nth-child(1) > app-search-company-search-filter > div.search-filter-wrap.t-text-primary-850.t-pb-5 > div.collapse-content.ng-star-inserted > app-search-company-general-information-filters > div.t-mb-4.pendo-GlPint4RqjBdpO6n2ifd > app-search-cognism-ng-select > div > ng-select > div > div > div.ng-input > input[type=text]')))
        for company in chunk:
            company_names_input.send_keys(company)
            company_names_input.send_keys(Keys.ENTER)

    def clear_companies(self):
        company_names_input = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                                                               '#app-container > app-prospects > section > div > app-search-company-search > div > div:nth-child(1) > app-search-company-search-filter > div.search-filter-wrap.t-text-primary-850.t-pb-5 > div.collapse-content.ng-star-inserted > app-search-company-general-information-filters > div.t-mb-4.pendo-GlPint4RqjBdpO6n2ifd > app-search-cognism-ng-select > div > ng-select > div > div > div.ng-input > input[type=text]')))
        company_names_input.click()
        for i in range(10):
            company_names_input.send_keys(Keys.BACKSPACE)

    def fetch_company_pages(self):
        if self.exact_match:
            self.toggle_exactmatch()
        self.set_countries()
        with open(f'company_pages_{self.country}_exactMatch={self.exact_match}', 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            for chunk in self.companies:
                self.set_companies(chunk)

                time.sleep(5)
                for info in self.driver.find_elements(By.CLASS_NAME, '_company-info'):
                    writer.writerow([info.find_element(By.XPATH, 'div[1]/div/a ').get_attribute("href")])
                self.clear_companies()

    def check_is_final_page(self):
        span_text = self.fetch_element(By.XPATH, '//*[@id="company-profile-container"]/section[2]/div/nz-tabset/div/div/div[1]/app-search-company-details-page-employees-section/div/div[2]/div[2]/app-paginationx/pagination-template/span').text.split(' ')
        return int(span_text[2]) == int(span_text[4])

    def set_role_filters(self):
        self.fetch_element(By.XPATH, '//*[@id="company-profile-container"]/section[2]/div/nz-tabset/div/div/div[1]/app-search-company-details-page-employees-section/div/div[1]/app-search-company-details-page-employees-filters/div/div/div[2]/button').click()
        self.fetch_element(By.CLASS_NAME, 'cdk-overlay-pane').find_element(By.TAG_NAME, 'label').click()

    def scrap_page(self, page_url):
        self.driver.get(page_url)
        company_name = self.fetch_element(By.XPATH, '//*[@id="company-profile-container"]/section[1]/div[1]/div[1]/div/div').text
        employees = []
        self.set_role_filters()
        time.sleep(3)
        while True:
            contacts_table = self.fetch_element(By.CSS_SELECTOR, '#company-profile-container > section.t-mx-4.t-mb-4._company-page-table.ng-star-inserted > div > nz-tabset > div > div > div.ant-tabs-tabpane.ant-tabs-tabpane-active.ng-star-inserted > app-search-company-details-page-employees-section > div > nz-table > nz-spin > div > div > nz-table-inner-default > div > table > tbody').find_elements(By.TAG_NAME, 'tr')
            for row in contacts_table:
                name = row.find_element(By.XPATH, 'td[2]/div/div/a').text
                title = row.find_element(By.XPATH, 'td[3]/div').text
                address = row.find_element(By.XPATH, 'td[5]/div').text
                numbers = []
                mail = ''
                try:
                    phone_tag = row.find_element(By.XPATH, 'td[4]/app-search-view-phone-email/div/div[1]/div/button')
                    phone_tag.click()
                    time.sleep(3)
                    numbers = [content.text.replace('\n', ' ') for content in row.find_element(By.XPATH,'td[4]/app-search-view-phone-email/div/div[1]/div/div/div').find_elements(By.TAG_NAME, 'div') if len(content.text) > 0]
                except:
                    numbers.append('error fetching number')
                    print('error fetching number')

                try:
                    row.find_element(By.XPATH, 'td[4]/app-search-view-phone-email/div/div[2]/div/button').click()
                    time.sleep(3)
                    mail = row.find_element(By.XPATH, 'td[4]/app-search-view-phone-email/div/div[2]/div/div/div/div').text
                except:
                    mail = 'No mail found'
                    print('No mail found')

                details = (company_name, name, title, mail, ' '.join(numbers), address)
                employees.append(details)

            if self.check_is_final_page():
                break
            self.fetch_element(By.XPATH, '//*[@id="company-profile-container"]/section[2]/div/nz-tabset/div/div/div[1]/app-search-company-details-page-employees-section/div/div[2]/div[2]/app-paginationx/pagination-template/div[2]/a').click()
            time.sleep(5)
        return employees

    def scrap_all_pages(self):
        with open(f'company_pages_{self.country}_exactMatch={self.exact_match}', 'r') as csv_file:
            with open(f'results_{self.country}_exactMatch={self.exact_match}.csv', 'a', newline='') as csvfile:
                writer = csv.writer(csvfile, delimiter=';')
                csv_reader = csv.reader(csv_file)
                for row in csv_reader:
                    try:
                        link = row[0]
                        employees = self.scrap_page(link)
                        for employee in employees:
                            writer.writerow(list(employee))
                    except Exception as e:
                        print('error for: ' + row[0])
                        print(e)


if __name__ == "__main__":
    # Initialize argparse
    parser = argparse.ArgumentParser(description='Fetch company pages and scrap all pages.')

    parser.add_argument('--country', type=str, default='France', help='Name of the country (default: France). KSA for Saudi Arabia')
    parser.add_argument('--p', action='store_true', default=False, help='True if you want to get company pages, False if you just want to scrap known pages')
    args = parser.parse_args()

    # Call read_file function with provided arguments
    country = args.country
    if country == 'KSA':
        country = 'Saudi Arabia'
    companies = read_file(target_file, input_sheet, country)


    # Initialize UserAgent with fetched companies and countries
    agent = UserAgent(email, pwd, companies, country)
    # Fetch company pages
    if args.p:
        agent.fetch_company_pages()

    # Scrap all pages
    agent.scrap_all_pages()

