from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


titles = ('Chief Marketing Officer', 'Chief Digital Officer', 'Content Manager', 'Digital Manager', 'Customer Experience Manager', 'Content Composer', 'Head of Marketing', 'Head of Digital', 'Head of Customer Experience' , 'Head of Digital Experience')

companies = ('AGL','Alinta Energy','ANZ Bank','Australia Post','Bank of Queensland','Bendigo Bank','BHP','CBA','Cenitex','Coles','Computershare','Crown','Dept of Defence','Dept of Health','Energy Aust','Frucor Suntory','GCP Asia Pacific','HCF','Just Group','Kmart','LeasePlan','Momentum Energy','Myer','NAB','NESA','Nufarm','NZ Police','Office Brands','Office Works','Orica','Origin Energy','QBE','QLD Dept of Transport','SA Health','SA Pathalology','Simply Energy','Stracco','Suncorp','The Good Guys/JB HiFi','The Star','Toll','Transport NSW','Westpac','Woolworths'
)

def main():
    driver = webdriver.Chrome()
    loginToLinkedInSalesNav(driver)
    startEmptySearchInSalesNav(driver)
    selectGeographyInSearch(driver, "Australia")
    selectTitlesInSearch(driver, titles)
    selectCompaniesInSearch(driver, companies)
    time.sleep(60)
    driver.close()


def loginToLinkedInSalesNav(driver):    
    
    driver.get("https://www.linkedin.com")

    login_form_pw = driver.find_element_by_id('session_password')
    login_form_id = driver.find_element_by_id('session_key')
    login_form_btn = driver.find_element_by_class_name("sign-in-form__submit-button")

    file_id = open('file_id.txt','r')
    id = file_id.read()
    file_id.close()

    file_password = open('file_password.txt','r')
    password = file_password.read()
    file_password.close()

    login_form_id.send_keys(id)
    login_form_pw.send_keys(password)
    login_form_btn.send_keys(Keys.RETURN)

    driver.get("https://www.linkedin.com/sales/homepage")


def startEmptySearchInSalesNav(driver):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.ID,'global-typeahead-search-input')))
    finally:
        search_bar = driver.find_element_by_id('global-typeahead-search-input');
        search_bar.send_keys(Keys.RETURN)


def selectGeographyInSearch(driver, country):

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


def selectTitleInSearch(driver, title):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[12]')))

    finally:
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]//div[@class="ph4 pb4"]/input')
        filter_search_bar.send_keys(title)
        filter_search_bar.send_keys(Keys.RETURN)
        

def selectTitlesInSearch(driver, titles):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[12]')))

    finally:
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]//div[@class="ph4 pb4"]/input')
        for title in titles:
            filter_search_bar.send_keys(title)
            filter_search_bar.send_keys(Keys.RETURN)


def selectCompaniesInSearch(driver, companies):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[7]')))

    finally:
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]//div[@class="ph4 pb4"]/input')
        for company in companies:
            filter_search_bar.send_keys(company)
            filter_search_bar.send_keys(Keys.RETURN)


main()
