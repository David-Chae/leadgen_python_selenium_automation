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


titles =[]
industry = []
companies = []
geographies = []

leads = []
page_urls = []

class Profile:

    def __init__(self, company, fullname, job_title, location, url):
        self._company = company
        self._job_title = job_title
        self._location = location
        self._url = url
        self._full_name = self.get_full_name(fullname)
        self._first_name = self.get_first_name(fullname)
        self._last_name = self.get_last_name(fullname)

    def get_company(self):
        return self._company
    
    def get_fullname(self):
        return self._fullname

    def get_job_title(self):
        return self._job_title

    def get_location(self):
        return self._location

    def get_url(self):
        return self._url

    def get_full_name(self, fullname):
        pattern = re.compile('[\W_]+', re.UNICODE)
        return re.sub(pattern, ' ', fullname)

    def get_first_name(self, fullname):
        name = HumanName(self.get_full_name(fullname))
        return name.first
    
    def get_last_name(self, fullname):
        name = HumanName(self.get_full_name(fullname))
        return name.last


def main():
    test_search()
    
    

def test_search():

    
    #Take record of time that this program started running.
    start_time = time.time()

    driver = webdriver.Chrome("./chromedriver.exe")

    #Log into Linkedin Sales Navigator.
    log_into_linked_in_sales_nav(driver)

    #Open an empty search page in Sales Navigator    
    start_empty_search_in_sales_nav(driver)

    #Following line of code has been commented out because the Linkedin returns too many request message.
    #select_companies_in_search(driver, company_list)
    company = "Trade Me"
    select_a_company_in_search(driver, company)

    #Select CXO as a seniority level. 
    select_seniority_in_search(driver, "CXO")
    select_seniority_in_search(driver, "VP")
    select_seniority_in_search(driver, "Director")
    select_seniority_in_search(driver, "Senior")
    select_seniority_in_search(driver, "Manager")

    #Search and then select Australia as geographical location of the leads.
    search_geography_in_search(driver, "Australia")
    select_geography_in_search(driver)
    time.sleep(1)
    search_geography_in_search(driver, "Indonesia")
    select_geography_in_search(driver)
    time.sleep(1)
    search_geography_in_search(driver, "Singapore")
    select_geography_in_search(driver)

    #Search and then select function of the leads.
    #search_function_in_search(driver, "Information Technology")
    #select_function_in_search(driver)

    #Search and then select retail as industry of the leads.
    #search_industry_in_search(driver, "retail")
    #select_industry_in_search(driver)

    #Select Chief Marketing Officer as a title in search.
    #for title in titles:
    #    select_title_in_search(driver, title)

    #driver.get("https://www.linkedin.com/sales/search/people?companySize=G%2CH%2CI&doFetchHeroCard=false&geoIncluded=101452733%2C102454443&industryIncluded=43%2C75%2C27%2C68%2C69%2C59%2C116%2C92%2C25%2C148%2C8&logHistory=true&rsLogId=1389016684&searchSessionId=COVLz%2FvhSsiVxYa0WgCntg%3D%3D&titleIncluded=Chief%2520Marketing%2520Officer%3A716%2CChief%2520Information%2520Officer%3A203%2CChief%2520Experience%2520Officer%3A30143%2CHead%2520Of%2520Customer%2520Experience%3A18497%2CChief%2520Digital%2520Officer%3A25884%2CChief%2520Digital%2520Transformation%2520Officer&titleTimeScope=CURRENT")
    
    
    #Zoom the browser to 60%.
    driver.execute_script("document.body.style.zoom='60%'")

    scroll_down(driver)

    #Get the number of pages in the search results. page_num is a string.
    page_num = get_num_of_search_result_pages(driver)
    print("You are now ready to move on to working with " + page_num + " pages.")
    
    #Get the number of search results in the current page.
    results_num = get_num_of_search_results_in_current_page(driver)
    print("You are now ready to move on to working with " + str(results_num) + " results in current page.")

    #Open each pages in the search. Append all page urls to page_urls list. 
    iterate_through_pages(driver)

    #Close the browser and its process to prevent out of memory issue.
    #driver.quit()
    
    #Open each page in the search one by one. 
    for url in page_urls:
        #Bring current page
        driver.get(url)
        #Go through all the search results in this page.
        iterate_through_results(driver)

    print("All results have been printed.")

    #Write the copied details into an excel file.
    filename = "ByteDance_" + company + ".xlsx"
    write_leads_to_excel_file(filename, company)
    print("All leads data have been written to xlsx file.")
    
    time.sleep(5)
    

    print("---This program took %s seconds ---" % (time.time() - start_time))    


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


def populate_geographies():
    with open('geographies.txt') as f:
        geographies_list = f.readlines()
        for geo in gepgraphies_list:
            geo = geo.rstrip("\n")
            geographies.append(seniority)

def populate_seniorities():
    with open('seniorities.txt') as f:
        seniorities_list = f.readlines()
        for seniority in seniorities_list:
            seniority = seniority.rstrip("\n")
            seniorities.append(seniority)


def populate_companies():
    with open('companies.txt') as f:
        comps = f.readlines()
        for comp in comps:
            comp = comp.rstrip("\n")
            companies.append(comp)


def populate_job_titles():
    with open('jobtitles.txt') as f:
        jobs = f.readlines()
        for job in jobs:
            job = job.rstrip("\n")
            titles.append(job)



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

    try:
        wait = WebDriverWait(driver, 10)
        # Wait until seniority tab element is found.
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[10]')))
        # Make seniority variable refer to Seniority level tab element in the Sales Navigator.
        seniority = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[10]')
        #Click open the Seniority level tab.
        seniority.click()

        lv = '"'+level+'"'

        # XPATH must be as follows. The . checks the whole string value of the button element
        # Explanation is at https://stackoverflow.com/questions/23676537/xpath-for-button-having-text-as-new
        # '//form[@class="search-filter__form"]/ul/li[10]/div/div/div/ol/li/button[contains(.,' + lv + ' )]'
        
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[10]/div/div/div/ol/li/button[contains(.,' + lv + ' )]')))
        # Get seniority list
        seniority = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[10]/div/div/div/ol/li/button[contains(.,' + lv + ' )]')
        seniority.click()

    except Exception as e:
        print(e)
        driver.quit()




def select_function_in_search(driver, category):

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[11]')))
    
        function = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]')
        function.click()

        function_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/input')
        function_search_bar.send_keys(category)
        function_search_bar.send_keys(Keys.RETURN)
    except StaleElementReferenceException:
        driver.refresh()
        select_function_in_search(driver, category)


    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button')))
    
        function_country_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button')
        function_country_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException):
        #driver.refresh()
        select_function_in_search(driver, category)




def search_industry_in_search(driver, industry):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[8]')))
        industry_filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[8]')
        industry_filter.click()
        industry_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[8]//div[@class="ph4 pb4"]/input')
        industry_search_bar.send_keys(industry)
        industry_search_bar.send_keys(Keys.RETURN)
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        search_industry_in_search(driver, industry)


def select_industry_in_search(driver):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[8]//div[@class="ph4 pb4"]/ol/li[1]/button')))
    
        industry_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[8]//div[@class="ph4 pb4"]/ol/li[1]/button')
        industry_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_industry_in_search(driver)




def search_function_in_search(driver, function):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[11]')))
        function_filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]')
        function_filter.click()
        function_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/input')
        function_search_bar.send_keys(function)
        function_search_bar.send_keys(Keys.RETURN)
        
    except (StaleElementReferenceException , TimeoutException) as e:
        driver.refresh()
        print(e)
        search_function_in_search(driver, function)
        


def select_function_in_search(driver):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button')))
    
        function_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button')
        function_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException) as e:
        driver.refresh()
        print(e)
        select_function_in_search(driver)




def search_geography_in_search(driver, country):

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[5]')))
        geography = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]')
        geography.click()
        geography_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/input')
        geography_search_bar.send_keys(country)
        geography_search_bar.send_keys(Keys.RETURN)
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        search_geography_in_search(driver, country)


def select_geography_in_search(driver):    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/ol/li[1]/button')))
    
        geography_country_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/ol/li[1]/button')
        geography_country_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_geography_in_search(driver)


def select_title_in_search(driver, title):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[12]')))
    
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]//div[@class="ph4 pb4"]/input')
        filter_search_bar.send_keys(title)
        filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        select_title_in_search(driver, title)
        driver.refresh()
        

def select_titles_in_search(driver, titles):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[12]')))
    
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]//div[@class="ph4 pb4"]/input')
        for title in titles:
            filter_search_bar.send_keys(title)
            filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_titles_in_search(driver, titles)


def select_companies_in_search(driver, companies):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[7]')))
    
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]//div[@class="ph4 pb4"]/input')
        for company in companies:
            filter_search_bar.send_keys(company)
            filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_companies_in_search(driver, companies)

def select_a_company_in_search(driver, company):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[7]')))
    
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]//div[@class="ph4 pb4"]/input')
        filter_search_bar.send_keys(company)
        filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_companies_in_search(driver, companies)

def get_num_of_search_result_pages(driver):
# I want to know the number of pages of search results   

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]')))
    
        page_num = driver.find_element(By.XPATH, '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]/li[last()]/button').text

    except NoSuchElementException:
        print("There is 1 page.")
        return "1"
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        get_num_of_search_result_pages(driver)

    return page_num

def get_num_of_search_results_in_current_page(driver):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]')))
        html_list = driver.find_elements(By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li')
        results_num = len(html_list)
    except StaleElementReferenceException:
        results_num = 0
    except TimeoutException:
        driver.refresh() 
        get_num_of_search_results_in_current_page(driver)
        
    return results_num




def iterate_through_pages(driver):
    
    page_num = int(get_num_of_search_result_pages(driver))
    curr = 1

    #Start populating a list of all page urls, starting with current page url.
    page_urls.append(driver.current_url)

    #iterate_through_results(driver)
    while curr < page_num:
        curr+=1
        
        try:
            wait = WebDriverWait(driver, 10)
            element = wait.until(EC.presence_of_element_located((By.XPATH, '//div/nav/ol[@class="search-results__pagination-list"]')))
            nextPage = driver.find_element(By.XPATH, '//div/nav/ol[@class="search-results__pagination-list"]/li[@class="selected cursor-pointer"]/following-sibling::li/button')
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

    try:
        #Wait until full name appears in DOM.
        elem1 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dt/a')))
        #Get the full name.
        fullname = driver.find_element(By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dt/a').text
        url = driver.find_element(By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dt/a').get_attribute('href')

    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
    except TimeoutException:
        fullname = "Timeout"
        url = "Timeout"
        
    try:
        #Wait until position appears in DOM.
        elem2 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[2]/span[1]')))
        #Get the position. 
        job_title = driver.find_element(By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[2]/span[1]').text

    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
    except TimeoutException:
        job_title = "Timeout"
        
        
    try:                               
        #Wait until position appears in DOM.
        elem3 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[2]/span[2]/span/a/span[1]')))
        #Get the position. 
        company = driver.find_element(By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[2]/span[2]/span/a/span[1]').text
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
    except TimeoutException:
        company = "Timeout"

    try:       
        #Wait until position appears in DOM.
        elem4 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[4]/ul/li')))
        #Get the position. 
        location = driver.find_element(By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[4]/ul/li').text
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

#main()
populate_job_titles()
