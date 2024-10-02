import time

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd


def main():
    path_to_chromedriver = 'C:/Program Files/Google/chromedriver.exe'
    path_to_excel = "assets/Book1.xlsx"

    df = pd.read_excel(path_to_excel)

    service = Service(executable_path=path_to_chromedriver)
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 10)
    form_url = "https://support.google.com/google-ads/troubleshooter/6099627?hl=en#ts=10634174"

    for index, row in df.iterrows():
        driver.get(form_url)

        contact_name_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='Contact name']")))
        contact_name_input.send_keys(row['Ad acct'])
        # send alt + enter
        contact_name_input.send_keys(Keys.ALT, Keys.ENTER)

        customer_company_name_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='End Customer Company Name']")))
        customer_company_name_input.send_keys(row['Company name'])

        address_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='Address']")))
        address_input.send_keys(row['adress'])

        country_code_dropdown = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.sc-select")))
        country_code_dropdown.click()
        time.sleep(20)  # TODO: Replace with wait for dropdown to open
        selected_option = wait.until(EC.element_to_be_clickable((By.XPATH, '//li[contains(text(), "United States (+1)")]')))
        selected_option.click()

        phone_number_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='Please provide a phone number we can call to reach you']")))
        phone_number_input.send_keys(row['phone'])

        contact_email_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='Contact email']")))
        contact_email_input.send_keys(row['Contact email'])

        email_cc_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='Email CC']")))
        email_cc_input.send_keys("")

        website_url_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='Website URL(s)']")))
        website_url_input.send_keys(row['Domain'])

        google_ads_customer_id_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-placeholder='Google Ads Customer ID']")))
        google_ads_customer_id_input.send_keys(row['CID'])

        link_to_page_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='Link to page on G2 site verifying Health Insurance Certification.']")))
        link_to_page_input.send_keys(row['g2 risk domain'])

        issue_summary_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "textarea[aria-label='Summary of the issue']")))
        issue_summary_input.send_keys("")

        time.sleep(10)

        submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Submit']")))
        submit_button.click()

        driver.delete_all_cookies()
        driver.refresh()
    driver.quit()


if __name__ == '__main__':
    main()
