from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def main():
    navigator.loginToLinkedInSalesNav()
    navigator.startEmptySearchInSalesNav()
    navigator.selectCountryInSearch("Australia")


def loginToLinkedInSalesNav():    
    
    driver.get("https://www.linkedin.com")

    login_form_pw = driver.find_element_by_id('session_password')
    login_form_id = driver.find_element_by_id('session_key')
    login_form_btn = driver.find_element_by_class_name("sign-in-form__submit-button")

    file = open('file.txt','r')
    content = file.read()
    file.close()

    login_form_id.send_keys("aksrud859@hanmail.net")
    login_form_pw.send_keys(content)
    login_form_btn.send_keys(Keys.RETURN)

    driver.get("https://www.linkedin.com/sales/homepage")


def startEmptySearchInSalesNav():
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.ID,'global-typeahead-search-input')))
    finally:
        search_bar = driver.find_element_by_id('global-typeahead-search-input');
        search_bar.send_keys(Keys.RETURN)

def selectCountryInSearch(country):

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[5]')))

    finally:
        geography = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]')
        geography.click()

        geography_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/input')
        geography_search_bar.send_keys(country)
        geography_search_bar.send_keys(Keys.RETURN)


    try:
        wait = WebDriverWait(driver, 1)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/ol/li[1]/button')))

    finally:
        geography_country_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/ol/li[1]/button')
        geography_country_btn.send_keys(Keys.RETURN)


#try:
#    wait = WebDriverWait(driver, 10)
#    element = wait.until(EC.presence_of_element_located((By.XPATH,'')))
#finally:
#    australia = driver.find_element(By.XPATH, '')
#    australia.click()

driver = webdriver.Chrome()
main()


time.sleep(60)
driver.close()
