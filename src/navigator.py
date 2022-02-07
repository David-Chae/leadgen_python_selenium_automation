from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
from nameparser import HumanName
import xlsxwriter
import time
import re, string
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from specifier import Profile , Account

job_titles_list =[]
industry = []
companies_list = []
geographies_list = []
seniorities_list = []
page_urls_list = []

leads = []
#This list contain the Account objects instantiated from going through results in account search results.
accounts = []
page_urls = []


def main():
    #leads_search()
    #accounts_search()
    extractFromAccountPages()

def extractFromAccountPages():
    #Take record of time that this program started running.
    start_time = time.time()

    #Read text file for opening compant account pages in Sales Navigator.
    populate_accounts_pages()

    #Instantiate a Chrome webdriver
    driver = webdriver.Chrome("./chromedriver.exe")
    
    #Log into Linkedin Sales Navigator.
    log_into_linked_in_sales_nav(driver)

    #Open an empty search page in Sales Navigator    
    start_empty_search_in_sales_nav(driver)

    for page_url in page_urls_list:
        driver.get(page_url)
        
        proceed = click_decision_makers(driver)

        if proceed == False: continue

        #Scroll Down
        scroll_down(driver)

        #Get the number of pages in the search results. page_num is a string.
        page_num = get_num_of_search_result_pages(driver)
        print("You are now ready to move on to working with " + str(page_num) + " pages.")
        
        #Get the number of search results in the current page.
        results_num = get_num_of_search_results_in_current_page(driver)
        print("You are now ready to move on to working with " + str(results_num) + " results in current page.")

        page_urls.clear()

        #Open each pages in the search. Append all page urls to page_urls list. 
        iterate_through_pages(driver)
        
        #Open each page in the search one by one. 
        for url in page_urls:
            #Bring current page
            driver.get(url)
            #Go through all the search results in this page.
            iterate_through_results(driver)

            
    #Write the copied details into an excel file.
    filename = "CX_decision_makers.xlsx"
    write_leads_to_excel_file(filename, "CX_Decision_makers")
    print("All accounts data have been written to xlsx file.")
    
    time.sleep(1)
    
    print("---This program took %s seconds ---" % (time.time() - start_time)) 
        
def click_decision_makers(driver):
    #Xpath needed for navigating to decision makers.
    btn_xpath = '//span[@class="account-top-card__search-leads"]/a'

    #GET Decision makers btn and Click it.
    try:
        #Wait until element appears in DOM.
        wait = WebDriverWait(driver, 20)
        element = wait.until(EC.element_to_be_clickable((By.XPATH, btn_xpath)))
        #Get the full name.
        btn = driver.find_element(By.XPATH, btn_xpath)
        btn.click()
                                                                  
    except StaleElementReferenceException:
        print("StaleElementReferenceException at clicking Decision makers button: ")
        print(driver.current_url)
        return False
    except TimeoutException:
        print("TimeoutException at clicking Decision makers button: ")
        print(driver.current_url)
        return False

    return True


def leads_search():
    #Take record of time that this program started running.
    start_time = time.time()

    #Read text files for selecting filters in Sales Navigator.
    #populate_geographies()
    #populate_seniorities()
    populate_companies()
    #populate_job_titles()

    #Instantiate a Chrome webdriver
    driver = webdriver.Chrome("./chromedriver.exe")

    #Log into Linkedin Sales Navigator.
    log_into_linked_in_sales_nav(driver)

    #Open an empty search page in Sales Navigator    
    start_empty_search_in_sales_nav(driver)

    #Select CXO as a seniority level.
    #for seniority in seniorities_list:
    #    select_seniority_in_search(driver, seniority)

    #Give time for Linkedin Sales Navigator to not think that I am Bot.
    #time.sleep(1)
    
    #Search and then select Australia as geographical location of the leads.
    #for geography in geographies_list:
    #    search_geography_in_search(driver, geography)
    #    select_geography_in_search(driver)

    #Select Chief Marketing Officer as a title in search.
    #for title in job_titles_list:
    #    select_title_in_search(driver, title)

    driver.get('https://www.linkedin.com/sales/search/people?_ntb=3EfQ5PAaR%2B2ekXE8mehYVQ%3D%3D&doFetchHeroCard=false&geoIncluded=101452733&keywords=omni%20channel%20solution&logHistory=true&rsLogId=1437635828&searchSessionId=muAfrqFmQMKH5JAHXMd6Sw%3D%3D&seniorityIncluded=8%2C6%2C7')
    
    #Following line of code has been commented out because the Linkedin returns too many request message.
    #select_companies_in_search(driver, companies_list)

    print("Current search results page URL: " + driver.current_url)

    #Zoom the browser to 60%.
    driver.execute_script("document.body.style.zoom='60%'")
    scroll_down(driver)

    #Get the number of pages in the search results. page_num is a string.
    page_num = get_num_of_search_result_pages(driver)
    print("You are now ready to move on to working with " + str(page_num) + " pages.")
    
    #Get the number of search results in the current page.
    results_num = get_num_of_search_results_in_current_page(driver)
    print("You are now ready to move on to working with " + str(results_num) + " results in current page.")

    #Open each pages in the search. Append all page urls to page_urls list. 
    iterate_through_pages(driver)
    
    #Open each page in the search one by one. 
    for url in page_urls:
        #Bring current page
        driver.get(url)
        #Go through all the search results in this page.
        iterate_through_results(driver)

    #Write the copied details into an excel file.
    filename = "CX_Leads.xlsx"
    write_leads_to_excel_file(filename, "CX_Ecommerce")
    print("All accounts data have been written to xlsx file.")
    
    time.sleep(1)
    
    print("---This program took %s seconds ---" % (time.time() - start_time)) 

def accounts_search():
    
    #Take record of time that this program started running.
    start_time = time.time()

    #Read text files for selecting filters in Sales Navigator.
    populate_geographies()
    populate_seniorities()
    populate_companies()
    populate_job_titles()

    #Instantiate a Chrome webdriver
    driver = webdriver.Chrome("./chromedriver.exe")

    #Log into Linkedin Sales Navigator.
    log_into_linked_in_sales_nav(driver)

    #Open an empty search page in Sales Navigator    
    start_empty_search_in_sales_nav(driver)

    #Select CXO as a seniority level.
    #for seniority in seniorities_list:
    #    select_seniority_in_search(driver, seniority)

    #Search and then select Australia as geographical location of the leads.
    #for geography in geographies_list:
    #    search_geography_in_search(driver, geography)
    #    select_geography_in_search(driver)
    #    time.sleep(1)

    #Select Chief Marketing Officer as a title in search.
    #for title in job_titles_list:
    #    select_title_in_search(driver, title)
    #time.sleep(10)
    
    #Following line of code has been commented out because the Linkedin returns too many request message.
    #select_companies_in_search(driver, companies_list)

    #Search and then select function of the leads.
    #search_function_in_search(driver, "Information Technology")
    #select_function_in_search(driver)

    #Search and then select retail as industry of the leads.
    #search_industry_in_search(driver, "retail")
    #select_industry_in_search(driver)

    driver.get('https://www.linkedin.com/sales/search/company?_ntb=3EfQ5PAaR%2B2ekXE8mehYVQ%3D%3D&geoIncluded=101452733&keywords=omni%20channel%20solution&searchSessionId=muAfrqFmQMKH5JAHXMd6Sw%3D%3D')
    #Zoom the browser to 60%.
    driver.execute_script("document.body.style.zoom='60%'")

    scroll_down(driver)

    #Get the number of pages in the search results. page_num is a string.
    page_num = get_num_of_search_result_pages(driver)
    print("You are now ready to move on to working with " + str(page_num) + " pages.")
    
    #Get the number of search results in the current page.
    results_num = get_num_of_search_results_in_current_page(driver)
    print("You are now ready to move on to working with " + str(results_num) + " results in current page.")

    #Open each pages in the search. Append all page urls to page_urls list. 
    iterate_through_pages(driver)
    
    #Open each page in the search one by one. 
    for url in page_urls:
        #Bring current page
        driver.get(url)
        #Go through all the search results in this page.
        iterate_through_companies(driver)

    print("All urls have been printed.")

    #Write the copied details into an excel file.
    filename = "Omni_Channel_Accounts.xlsx"
    write_accounts_to_excel_file(filename, "Test")
    print("All accounts data have been written to xlsx file.")
    
    time.sleep(1)
    

    print("---This program took %s seconds ---" % (time.time() - start_time))    


def iterate_through_companies(driver):
    results_num = get_num_of_search_results_in_current_page(driver)
    scroll_down(driver)

    if results_num > 0:
        curr = 1
        while curr <= results_num:
            get_company_data_from_search_result(driver, curr)
            curr+=1
    else:
        curr = 0


def get_company_data_from_search_result(driver, curr):
    #Get the number of results in the current page.
    results_num = get_num_of_search_results_in_current_page(driver)

    #Initialise this WebDriverWait instance so I can use in the loop below.
    wait = WebDriverWait(driver, 10)

    #Use string of pointer for XPATH
    pointer_str = str(curr)

    #Xpaths for fullname, job title, company and location.
    company_name_xpath = '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']//dt[@class="result-lockup__name"]/a'
    industry_xpath = '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']//li[1][@class="result-lockup__misc-item"]'
    headcount_xpath = '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']//li[2][@class="result-lockup__misc-item"]'
    location_xpath = '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']//li[3][@class="result-lockup__misc-item"]'
    description_xpath = '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']//dd[@class="result-lockup__description"]/div/span[2]'

    #GET COMPANY NAME AND URL
    try:
        #Wait until element appears in DOM.
        elem1 = wait.until(EC.presence_of_element_located((By.XPATH, company_name_xpath)))
        #Get the full name.
        company_name = driver.find_element(By.XPATH, company_name_xpath).text
        company_name_url = driver.find_element(By.XPATH, company_name_xpath).get_attribute('href')
                                                          
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_company_data_from_search_result(driver, pointer)
    except TimeoutException:
        fullname = "Timeout"
        url = "Timeout"

    #GET INDUSTRY (CATEGORY)
    try:
        #Wait until element appears in DOM.
        elem2 = wait.until(EC.presence_of_element_located((By.XPATH, industry_xpath)))
        #Get the position. 
        industry = driver.find_element(By.XPATH, industry_xpath).text

    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_company_data_from_search_result(driver, pointer)
    except TimeoutException:
        industry = "Timeout"
        
    #GET HEADCOUNT
    try:                               
        #Wait until element appears in DOM.
        elem3 = wait.until(EC.presence_of_element_located((By.XPATH, headcount_xpath)))
        #Get the position. 
        headcount = driver.find_element(By.XPATH, headcount_xpath).text
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_company_data_from_search_result(driver, pointer)
    except TimeoutException:
        headcount = "Timeout"

    #GET LOCATION
    try:       
        #Wait until element appears in DOM.
        elem4 = wait.until(EC.presence_of_element_located((By.XPATH, location_xpath)))
        #Get the position. 
        location = driver.find_element(By.XPATH, location_xpath).text
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_company_data_from_search_result(driver, pointer)
    except TimeoutException:
        location = "Timeout"

    #GET COMPANY DESCRIPTION
    try:       
        #Wait until element appears in DOM.
        elem4 = wait.until(EC.presence_of_element_located((By.XPATH, description_xpath)))
        #Get the position. 
        description = driver.find_element(By.XPATH, description_xpath).text
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_company_data_from_search_result(driver, pointer)
    except TimeoutException:
        description = "Timeout"

    company = Account(company_name, company_name_url, industry, headcount, location, description)
    accounts.append(company)

    print(company._company_name + "%^&" + company._industry + "%^&" + company._headcount + "%^&" + company._location + "%^&" + company._description + "%^&" + company._company_name_url)
    print("  ")


def write_accounts_to_excel_file(file_name, sheet_name):
    # file_name e.g "leads_.xlsx"
    workbook = xlsxwriter.Workbook(file_name)
    # sheet_name e.g "HCL Appscan 2022"
    worksheet = workbook.add_worksheet(sheet_name)

    row = 1
    col = 0

    header = ["COMPANY NAME", "INDUSTRY", "HEADCOUNT", "LOCATION", "DESCRIPTION", "LINKEDIN URL"]

    for hd in header:
        worksheet.write(0, col, hd)
        col+=1

    for account in accounts:
        worksheet.write(row, 0, account._company_name)
        worksheet.write(row, 1, account._industry)
        worksheet.write(row, 2, account._headcount)
        worksheet.write(row, 3, account._location)
        worksheet.write(row, 4, account._description)
        worksheet.write(row, 5, account._company_name_url)
        print(str(row) + " accounts written to file.")
        row+=1        
    
    workbook.close()


def temp_search(url):

    #Make a temporary browser.
    temp_driver = webdriver.Chrome()

    #Log into Linkedin Sales Navigator.
    log_into_linked_in_sales_nav(temp_driver)

    #Bring current page
    temp_driver.get(url)
    
    #Go through all the search results in this page.
    iterate_through_results(temp_driver)
    #Close the browser and its process in the background.
    temp_driver.quit()

def populate_accounts_pages():
    with open('accounts_pages.txt') as f:
        pages_urls = f.readlines()
        for page in pages_urls:
            page = page.rstrip("\n")
            page_urls_list.append(page)

def populate_geographies():
    with open('geographies.txt') as f:
        geographies = f.readlines()
        for geo in geographies:
            geo = geo.rstrip("\n")
            geographies_list.append(geo)

def populate_seniorities():
    with open('seniorities.txt') as f:
        seniorities = f.readlines()
        for seniority in seniorities:
            seniority = seniority.rstrip("\n")
            seniorities_list.append(seniority)


def populate_companies():
    with open('companies.txt') as f:
        companies = f.readlines()
        for comp in companies:
            comp = comp.rstrip("\n")
            companies_list.append(comp)


def populate_job_titles():
    with open('jobtitles.txt') as f:
        jobs = f.readlines()
        for job in jobs:
            job = job.rstrip("\n")
            job_titles_list.append(job)



def log_into_linked_in_sales_nav(driver):    
    
    driver.get("https://www.linkedin.com")

    try:
        login_form_pw = driver.find_element_by_id('session_password')
        login_form_id = driver.find_element_by_id('session_key')
        login_form_btn = driver.find_element_by_class_name("sign-in-form__submit-button")
        
        file_id = open('file_id.txt','r')
        linkedin_id = file_id.read()
        file_id.close()

        file_password = open('file_password.txt','r')
        linkedin_password = file_password.read()
        file_password.close()

        login_form_id.send_keys(linkedin_id)
        login_form_pw.send_keys(linkedin_password)
        login_form_btn.send_keys(Keys.RETURN)

        driver.get("https://www.linkedin.com/sales/homepage")
        
    except StaleElementReferenceException:
        driver.refresh()
        log_into_linked_in_sales_nav(driver)


def start_empty_search_in_sales_nav(driver):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.ID,'global-typeahead-search-input')))
    
        search_bar = driver.find_element_by_id('global-typeahead-search-input');
        search_bar.send_keys(Keys.RETURN)

    except StaleElementReferenceException:
        driver.refresh()
        start_empty_search_in_sales_nav(driver)
        


def select_seniority_in_search(driver, level):
    #Xpaths needed for this function
    seniority_xpath = '//form[@class="search-filter__form"]/ul/li[10]'

    try:
        wait = WebDriverWait(driver, 10)
        # Wait until seniority tab element is found.
        element = wait.until(EC.presence_of_element_located((By.XPATH, seniority_xpath)))
        # Make seniority variable refer to Seniority level tab element in the Sales Navigator.
        seniority = driver.find_element(By.XPATH, seniority_xpath)
        #Click open the Seniority level tab.
        seniority.click()

        lv = '"'+level+'"'

        # XPATH must be as follows. The . checks the whole string value of the button element
        # Explanation is at https://stackoverflow.com/questions/23676537/xpath-for-button-having-text-as-new
        # '//form[@class="search-filter__form"]/ul/li[10]/div/div/div/ol/li/button[contains(.,' + lv + ' )]'
        seniority_btn_xpath = '//form[@class="search-filter__form"]/ul/li[10]/div/div/div/ol/li/button[contains(.,' + lv + ' )]'
        
        element = wait.until(EC.presence_of_element_located((By.XPATH, seniority_btn_xpath)))
        # Get seniority list
        seniority_btn = driver.find_element(By.XPATH, seniority_btn_xpath)
        seniority_btn.click()

    except Exception as e:
        print(e)
        driver.quit()


def select_function_in_search(driver, category):
    #Xpaths needed for this function
    function_xpath = '//form[@class="search-filter__form"]/ul/li[11]'
    function_search_bar_xpath = '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/input'
    function_btn_xpath = '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button'

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, function_xpath)))
    
        function = driver.find_element(By.XPATH, function_xpath)
        function.click()

        function_search_bar = driver.find_element(By.XPATH, function_search_bar_xpath)
        function_search_bar.send_keys(category)
        function_search_bar.send_keys(Keys.RETURN)
    except StaleElementReferenceException:
        driver.refresh()
        select_function_in_search(driver, category)


    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, function_btn_xpath)))
    
        function_btn = driver.find_element(By.XPATH, function_btn_xpath)
        function_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException):
        #driver.refresh()
        select_function_in_search(driver, category)


def search_industry_in_search(driver, industry):
    #Xpaths needed for this function
    industry_filter_xpath = '//form[@class="search-filter__form"]/ul/li[8]'
    industry_search_bar_xpath = '//form[@class="search-filter__form"]/ul/li[8]//div[@class="ph4 pb4"]/input'
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, industry_filter_xpath)))
        industry_filter = driver.find_element(By.XPATH, industry_filter_xpath)
        industry_filter.click()
        industry_search_bar = driver.find_element(By.XPATH, industry_search_bar_xpath)
        industry_search_bar.send_keys(industry)
        industry_search_bar.send_keys(Keys.RETURN)
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        search_industry_in_search(driver, industry)


def select_industry_in_search(driver):
    #Xpaths needed for this function
    industry_btn_xpath = '//form[@class="search-filter__form"]/ul/li[8]//div[@class="ph4 pb4"]/ol/li[1]/button'
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, industry_btn_xpath)))
    
        industry_btn = driver.find_element(By.XPATH, industry_btn_xpath)
        industry_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_industry_in_search(driver)
        

def search_function_in_search(driver, function):
    #Xpaths needed for this function
    function_filter_xpath = '//form[@class="search-filter__form"]/ul/li[11]'
    function_search_bar_xpath = '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/input'
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, function_filter_xpath)))
        function_filter = driver.find_element(By.XPATH, function_filter_xpath)
        function_filter.click()
        function_search_bar = driver.find_element(By.XPATH, function_search_bar_xpath)
        function_search_bar.send_keys(function)
        function_search_bar.send_keys(Keys.RETURN)
        
    except (StaleElementReferenceException , TimeoutException) as e:
        driver.refresh()
        print(e)
        search_function_in_search(driver, function)
        


def select_function_in_search(driver):
    #Xpath needed for this function
    function_btn_xpath = '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button'
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, function_btn_xpath)))
    
        function_btn = driver.find_element(By.XPATH, function_btn_xpath)
        function_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException) as e:
        driver.refresh()
        print(e)
        select_function_in_search(driver)


def search_geography_in_search(driver, country):
    #Xpath needed for this function
    geography_xpath  = '//form[@class="search-filter__form"]/ul/li[5]'
    geography_search_bar_xpath = '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/input'

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,geography_xpath)))
        geography = driver.find_element(By.XPATH, geography_xpath)
        geography.click()
        geography_search_bar = driver.find_element(By.XPATH, geography_search_bar_xpath)
        geography_search_bar.send_keys(country)
        geography_search_bar.send_keys(Keys.RETURN)
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        search_geography_in_search(driver, country)


def select_geography_in_search(driver):
    #Xpath needed for this function
    geography_country_btn_xpath = '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/ol/li[1]/button'
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, geography_country_btn_xpath)))
    
        geography_country_btn = driver.find_element(By.XPATH, geography_country_btn_xpath)
        geography_country_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_geography_in_search(driver)


def select_title_in_search(driver, title):
    #Xpath needed for this function
    title_filter_xpath = '//form[@class="search-filter__form"]/ul/li[12]'
    filter_search_bar_xpath = '//form[@class="search-filter__form"]/ul/li[12]//div[@class="ph4 pb4"]/input'
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, title_filter_xpath)))
    
        filter = driver.find_element(By.XPATH, title_filter_xpath)
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, filter_search_bar_xpath)
        filter_search_bar.send_keys(title)
        filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        select_title_in_search(driver, title)
        driver.refresh()
        

def select_titles_in_search(driver, titles):
    #Xpaths needed for this function
    title_filter_xpath = '//form[@class="search-filter__form"]/ul/li[12]'
    filter_search_bar_xpath = '//form[@class="search-filter__form"]/ul/li[12]//div[@class="ph4 pb4"]/input'
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, title_filter_xpath)))
    
        title_filter = driver.find_element(By.XPATH, title_filter_xpath)
        title_filter.click()

        filter_search_bar = driver.find_element(By.XPATH, filter_search_bar_xpath)

        for title in titles:
            filter_search_bar.send_keys(title)
            filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_titles_in_search(driver, titles)


def select_companies_in_search(driver, companies):
    #Xpaths needed for this function
    nav_filter_xpath = '//form[@class="search-filter__form"]/ul/li[7]'
    filter_search_bar_xpath = '//form[@class="search-filter__form"]/ul/li[7]//div[@class="ph4 pb4"]/input'
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.element_to_be_clickable((By.XPATH, nav_filter_xpath)))
    
        nav_filter = driver.find_element(By.XPATH, nav_filter_xpath)
        nav_filter.click()

        filter_search_bar = driver.find_element(By.XPATH, filter_search_bar_xpath)
        
        for company in companies:
            filter_search_bar.send_keys(company)
            filter_search_bar.send_keys(Keys.RETURN)
            time.sleep(1)
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_companies_in_search(driver, companies)

def select_a_company_in_search(driver, company):
    #Xpaths needed for this function.
    nav_filter_xpath = '//form[@class="search-filter__form"]/ul/li[7]'
    filter_search_bar_xpath = '//form[@class="search-filter__form"]/ul/li[7]//div[@class="ph4 pb4"]/input'
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,nav_filter_xpath)))
    
        nav_filter = driver.find_element(By.XPATH, nav_filter_xpath)
        nav_filter.click()

        filter_search_bar = driver.find_element(By.XPATH, filter_search_bar_xpath)
        filter_search_bar.send_keys(company)
        filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_companies_in_search(driver, companies)

def get_num_of_search_result_pages(driver):
    # I want to know the number of pages of search results   
    # Xpaths needed fot this function
    pagenation_list_xpath = '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]'
    page_number_xpath = '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]/li[last()]/button'
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, pagenation_list_xpath)))
        page_num = driver.find_element(By.XPATH, page_number_xpath).text
        page_num = int(page_num)

    except NoSuchElementException:
        print("There is 1 page.")
        return 1
        
    except (StaleElementReferenceException , TimeoutException):
        return 0

    return page_num

def get_num_of_search_results_in_current_page(driver):

    #Xpaths needed for this function.
    search_results_xpath = '//section[@id="results"]//ol[@class="search-results__result-list"]'
    html_list_xpath = '//section[@id="results"]//ol[@class="search-results__result-list"]/li'
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, search_results_xpath)))
        html_list = driver.find_elements(By.XPATH, html_list_xpath)
        results_num = int(len(html_list))
    except StaleElementReferenceException:
        results_num = 0
    except TimeoutException:
        results_num = 0
        
    return results_num





def iterate_through_pages(driver):
    
    page_num = get_num_of_search_result_pages(driver)
    curr = 1

    #Xpaths needed for iterating through the pages.
    page_list_xpath = '//div/nav/ol[@class="search-results__pagination-list"]'
    next_page_xpath = '//div/nav/ol[@class="search-results__pagination-list"]/li[@class="selected cursor-pointer"]/following-sibling::li/button'

    #Start populating a list of all page urls, starting with current page url.
    page_urls.append(driver.current_url)

    #iterate_through_results(driver)
    while curr < page_num:
        curr+=1
        
        try:
            wait = WebDriverWait(driver, 10)
            element = wait.until(EC.presence_of_element_located((By.XPATH, page_list_xpath )))
            nextPage = driver.find_element(By.XPATH, next_page_xpath)
            nextPage.send_keys(Keys.RETURN)
            
            time.sleep(2)
            page_urls.append(driver.current_url)
        except (StaleElementReferenceException , TimeoutException):
            curr-=1
            driver.refresh()


def iterate_through_results(driver):
    results_num = get_num_of_search_results_in_current_page(driver)
    scroll_down(driver)

    if results_num > 0:
        curr = 1
        while curr <= results_num:
            #open_search_results(driver, curr)
            get_profile_data_from_search_result(driver, curr)
            curr+=1
    else:
        curr = 0


# I want to go through the results in the page and print one by one on console.
# This function assumes that the driver is currently at a sales navigator page with results showing from search.
# For each result, this function grab profile data: full name, first name, last name, location, position, company and url of LinkedIn profile.
# This populates leads list variable so it can be written to an excel file.
def get_profile_data_from_search_result(driver, pointer):
    
    #Get the number of results in the current page.
    results_num = get_num_of_search_results_in_current_page(driver)

    #Initialise this WebDriverWait instance so I can use in the loop below.
    wait = WebDriverWait(driver, 10)

    #Use string of pointer for XPATH
    pointer_str = str(pointer)

    #Xpaths for fullname, job title, company and location.
    fullname_xpath = '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']//dt[@class="result-lockup__name"]/a'
    jobtitle_xpath = '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']//dd[@class="result-lockup__highlight-keyword"]/span[1]'
    company_xpath = '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']//dd[@class="result-lockup__highlight-keyword"]/span[2]/span/a/span[1]'
    location_xpath = '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']//li[@class="result-lockup__misc-item"]'

    try:
        #Wait until full name appears in DOM.
        elem1 = wait.until(EC.presence_of_element_located((By.XPATH, fullname_xpath)))
        #Get the full name.
        fullname = driver.find_element(By.XPATH, fullname_xpath).text
        url = driver.find_element(By.XPATH, fullname_xpath).get_attribute('href')
                                                          
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
    except TimeoutException:
        fullname = "Timeout"
        url = "Timeout"
        
    try:
        #Wait until position appears in DOM.
        elem2 = wait.until(EC.presence_of_element_located((By.XPATH, jobtitle_xpath)))
        #Get the position. 
        job_title = driver.find_element(By.XPATH, jobtitle_xpath).text

    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
    except TimeoutException:
        job_title = "Timeout"
        
        
    try:                               
        #Wait until position appears in DOM.
        elem3 = wait.until(EC.presence_of_element_located((By.XPATH, company_xpath)))
        #Get the position. 
        company = driver.find_element(By.XPATH, company_xpath).text
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
    except TimeoutException:
        company = "Timeout"

    try:       
        #Wait until position appears in DOM.
        elem4 = wait.until(EC.presence_of_element_located((By.XPATH, location_xpath)))
        #Get the position. 
        location = driver.find_element(By.XPATH, location_xpath).text
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
    except TimeoutException:
        location = "Timeout"

    person = Profile(company, fullname, job_title, location, url)
    leads.append(person)

    print(person._company + "%^&" + person._full_name + "%^&" + person._first_name + "%^&" + person._last_name + "%^&" + person._location + "%^&" + person._job_title + "%^&" + person._url)
    print("  ")



def open_search_results(driver, curr):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li['+ str(curr) + ']/div[2]/div/div/div/article/section[@class="result-lockup"]/div/div/dl/dt[@class="result-lockup__name"]/a')))

        url = driver.find_element(By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li['+ str(curr) + ']/div[2]/div/div/div/article/section[@class="result-lockup"]/div/div/dl/dt[@class="result-lockup__name"]/a').get_attribute('href')
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[1])
        driver.get(url);
        
        grab_details(driver)

        time.sleep(2)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        open_search_results(driver, curr)


def grab_details(driver):

    driver.execute_script("document.body.style.zoom='60%'")

    try:
        wait = WebDriverWait(driver, 10)
        
        elem_fullname = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dt/span')))
        fullname = driver.find_element(By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dt/span').text

        #This if statement handles a case where clicking a link brings up locked Linkedin profile.
        if fullname != "LinkedIn Member":
        
            elem_location = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dd[@class="mt4 mb0"]/div')))
            elem_position = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dt')))
            elem_company = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dd[1]/span[2]')))
            
            job_title = driver.find_element(By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dt').text
            
            location = driver.find_element(By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dd[@class="mt4 mb0"]/div').text     
            company = driver.find_element(By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dd[1]/span[2]').text  
            url = driver.current_url

            person = Profile(company, fullname, job_title, location, url)
            leads.append(person)

            print(person._company + "%^&" + person._full_name + "%^&" + person._first_name + "%^&" + person._last_name + "%^&" + person._location + "%^&" + person._job_title + "%^&" + person._url)
            print("  ")
    
        else:
            print(driver.current_url)
            print("This is a locked Linkedin Member in Sales Navigator.")

    except StaleElementReferenceException:
        driver.refresh()
        grab_details(driver)


    

def scroll_down(driver):
    
    height = driver.execute_script("return document.documentElement.scrollHeight")

    #Divide the scroll height into 6 equal sections so driver can stop at each section.
    sect1 = height/6
    sect2 = sect1 * 2
    sect3 = sect1 * 3
    sect4 = sect1 * 4
    sect5 = sect1 * 5

    #Stop at each sections ending with the bottom of the page. Make sure all results load on the page.    
    driver.execute_script("window.scrollTo(0, " + str(sect1) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(sect2) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(sect3) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(sect4) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(sect5) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(height) + ");")
    time.sleep(1)

def write_leads_to_excel_file(file_name, sheet_name):
    # file_name e.g "leads_.xlsx"
    workbook = xlsxwriter.Workbook(file_name)
    # sheet_name e.g "HCL Appscan 2022"
    worksheet = workbook.add_worksheet(sheet_name)

    row = 1
    col = 0

    header = ["COMPANY", "FULL NAME", "FIRST NAME", "LAST NAME", "JOB TITLE", "LOCATION", "LINKEDIN URL"]

    for hd in header:
        worksheet.write(0, col, hd)
        col+=1

    for lead in leads:
        worksheet.write(row, 0, lead._company)
        worksheet.write(row, 1, lead._full_name)
        worksheet.write(row, 2, lead._first_name)
        worksheet.write(row, 3, lead._last_name)
        worksheet.write(row, 4, lead._job_title)
        worksheet.write(row, 5, lead._location)
        worksheet.write(row, 6, lead._url)
        print(str(row) + " leads written to file.")
        row+=1        
    
    workbook.close()

main()
