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


titles = ('Chief Marketing Officer')
#, 'Chief Digital Officer', 'Content Manager', 'Digital Manager', 'Customer Experience Manager', 'Content Composer', 'Head of Marketing', 'Head of Digital', 'Head of Customer Experience' , 'Head of Digital Experience')

companies = ('AGL','Alinta Energy','ANZ Bank','Australia Post','Bank of Queensland','Bendigo Bank','BHP','CBA','Cenitex','Coles','Computershare','Crown','Dept of Defence','Dept of Health','Energy Aust','Frucor Suntory','GCP Asia Pacific','HCF','Just Group','Kmart','LeasePlan','Momentum Energy','Myer','NAB','NESA','Nufarm','NZ Police','Office Brands','Office Works','Orica','Origin Energy','QBE','QLD Dept of Transport','SA Health','SA Pathalology','Simply Energy','Stracco','Suncorp','The Good Guys/JB HiFi','The Star','Toll','Transport NSW','Westpac','Woolworths'
)

leads = []



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

    driver = webdriver.Chrome()

    #Log into Linkedin Sales Navigator.
    log_into_linked_in_sales_nav(driver)

    #Open an empty search page in Sales Navigator    
    start_empty_search_in_sales_nav(driver)

    #Search and then select Australia as geographical location of the leads.
    search_geography_in_search(driver, "Australia")
    select_geography_in_search(driver)
    #search_geography_in_search(driver, "New Zealand")
    #select_geography_in_search(driver)

    #Select Chief Marketing Officer as a title in search.
    select_title_in_search(driver, 'Chief Marketing Officer')

    #Select Arts and Design as function in search.
    select_function_in_search(driver, "Arts and Design")
    
    #Zoom the browser to 60%.
    driver.execute_script("document.body.style.zoom='60%'")

    #Following line of code has been commented out because the Linkedin returns too many request message.
    #select_companies_in_search(driver, companies)

    #Get the number of pages in the search results. page_num is a string.
    page_num = get_num_of_search_result_pages(driver)
    print("You are now ready to move on to working with " + page_num + " pages.")
    
    #Get the number of search results in the current page.
    results_num = get_num_of_search_results_in_current_page(driver)
    print("You are now ready to move on to working with " + str(results_num) + " results in current page.")

    #Open each results in the search results and copy details into an object. Move to next page if necessary.
    iterate_through_pages(driver)
    print("All results have been printed.")

    #Write the copied details into an excel file.
    write_leads_to_excel_file("leads.xlsx", "Australia_CMO_Arts_N_Design")
    print("All leads data have been written to xlsx file.")
    
    time.sleep(6)
    driver.quit()

    print("---This program took %s seconds ---" % (time.time() - start_time))    






def log_into_linked_in_sales_nav(driver):    
    
    driver.get("https://www.linkedin.com")

    try:
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
        driver.refresh()
        select_function_in_search(driver, category)


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
        select_geography_in_search(driver, country)


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

def get_num_of_search_result_pages(driver):
# I want to know the number of pages of search results
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]')))
    
        page_num = driver.find_element(By.XPATH, '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]/li[last()]/button').text
        
        print("There are " + page_num + " pages")
        
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
        print("There are " + str(results_num) + " results in current page.")
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        get_num_of_search_results_in_current_page(driver)
        
    return results_num
    

def iterate_through_pages(driver):
    
    page_num = int(get_num_of_search_result_pages(driver))
    curr = 1

    iterate_through_results(driver)

    while curr < page_num:
        curr+=1

        try:
            wait = WebDriverWait(driver, 10)
            element = wait.until(EC.presence_of_element_located((By.XPATH, '//div/nav/ol[@class="search-results__pagination-list"]')))
        
            nextPage = driver.find_element(By.XPATH, '//div/nav/ol[@class="search-results__pagination-list"]/li[@class="selected cursor-pointer"]/following-sibling::li/button')
            nextPage.send_keys(Keys.RETURN)
            time.sleep(2)
            iterate_through_results(driver)
        except (StaleElementReferenceException , TimeoutException):
            curr-=1
            driver.refresh()


def iterate_through_results(driver):
    results_num = get_num_of_search_results_in_current_page(driver)
    scroll_down(driver)
    
    curr = 1
    while curr <= results_num:
        open_search_results(driver, curr)
        curr+=1


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
        elem_location = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dd[@class="mt4 mb0"]/div')))
        elem_position = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dt')))
        elem_company = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dd[1]/span[2]')))
        
        job_title = driver.find_element(By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dt').text
        fullname = driver.find_element(By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dt/span').text
        location = driver.find_element(By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dd[@class="mt4 mb0"]/div').text     
        company = driver.find_element(By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dd[1]/span[2]').text  
        url = driver.current_url

        person = Profile(company, fullname, job_title, location, url)
        leads.append(person)

        print("Company: " + person._company)
        print("Fullname: " + person._full_name)
        print("First name: " + person._first_name)
        print("Last name: " + person._last_name)
        print("Location: " +person._location)
        print("Job Title: " + person._job_title)
        print("Linkedin URL: " + person._url)

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
