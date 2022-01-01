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


#REQUIRED GLOBAL VARIABLES
leads = []
page_urls = []

#For quick shortcut in case I need to skip filter selection.
saved_search = "https://www.linkedin.com/sales/search/people?query=(recentSearchParam%3A(id%3A1352491996%2CdoLogHistory%3Atrue)%2Cfilters%3AList((type%3AREGION%2Cvalues%3AList((id%3A101452733%2Ctext%3AAustralia%2CselectionType%3AINCLUDED)))%2C(type%3AFUNCTION%2Cvalues%3AList((id%3A13%2Ctext%3AInformation%2520Technology%2CselectionType%3AINCLUDED)))%2C(type%3ATITLE%2Cvalues%3AList((id%3A5134%2Ctext%3AChief%2520Information%2520Security%2520Officer%2CselectionType%3AINCLUDED))%2CselectedSubFilter%3ACURRENT)%2C(type%3ACOMPANY_HEADCOUNT%2Cvalues%3AList((id%3AI%2Ctext%3A10%252C000%252B%2CselectionType%3AINCLUDED)))))&sessionId=x%2B7KGzm4R1mfkW%2FCYWCR4g%3D%3D"

#Profile class is for creating profile object.
#It is used to retrieve information about a person.
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


#This function specifies the kind of search to be done.
def test_search():
    #Take record of time that this program started running.
    start_time = time.time()

    driver = webdriver.Chrome()

    #Log into Linkedin Sales Navigator.
    log_into_linked_in_sales_nav(driver)
    
    #Open an empty search page in Sales Navigator
    start_empty_search_in_sales_nav(driver)
  
    #driver.get(saved_search)

    #Expand and collapse functions work.
    expand_collapse_filters(driver)
    time.sleep(2)

    specify_a_seniority(driver, "CXO")
    time.sleep(2)

    expand_collapse_filters(driver)
    
    #Zoom the browser to 60%.
    #driver.execute_script("document.body.style.zoom='60%'")

    #scroll_down(driver)

    #Get the number of pages in the search results. page_num is a string.
    #page_num = get_num_of_search_result_pages(driver)
    #print("You are now ready to move on to working with " + page_num + " pages.")

    #Get the number of search results in the current page.
    #results_num = get_num_of_search_results_in_current_page(driver)
    #print("You are now ready to move on to working with " + str(results_num) + " results in current page.")

    #Open each pages in the search. Append all page urls to page_urls list. 
    #iterate_through_pages(driver, int(page_num))

    #Close the browser and its process to prevent out of memory issue.
    #driver.quit()

    #Open each page in the search one by one. 
    #for url in page_urls:
        #Open all results in current url one by one. Grab details and append it to leads list.
    #    temp_search(url)

    #print("All results have been printed.")

    #Write the copied details into an excel file.
    #write_leads_to_excel_file("update_test.xlsx", "CISO")
    #print("All leads data have been written to xlsx file.")
    
    time.sleep(2)

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


def iterate_through_results(driver):
    results_num = get_num_of_search_results_in_current_page(driver)
    scroll_down(driver)
    
    curr = 1
    while curr <= results_num:
        element = driver.find_element_by_xpath('//ol[@class="artdeco-list background-color-white _border-search-results_1igybl"]/li[' + str(curr) + ']')
        driver.execute_script("return arguments[0].scrollIntoView();", element)
        get_profile_data_from_search_result(driver, curr)
        curr+=1

# I want to go through the results in the page and print one by one on console.
# This function assumes that the driver is currently at a sales navigator page with results showing from search.
# For each result, this function grab profile data: full name, first name, last name, location, position, company and url of LinkedIn profile.
# This populates leads list variable so it can be written to an excel file.
def get_profile_data_from_search_result(driver, pointer):
    
    #Get the number of results in the current page.
    results_num = get_num_of_search_results_in_current_page(driver)

    #Initialise this WebDriverWait instance so I can use in the loop below.
    wait = WebDriverWait(driver, 3)

    #Use string of pointer for XPATH
    pointer_str = str(pointer)
    
    try:
        #Wait until full name appears in DOM.
        elem1 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="artdeco-list background-color-white _border-search-results_1igybl"]/li[' + pointer_str + ']/div/div/div[2]/div[1]/div[1]/div/div[2]/div[1]/a')))
        #Get the full name.
        fullname = driver.find_element(By.XPATH, '//ol[@class="artdeco-list background-color-white _border-search-results_1igybl"]/li[' + pointer_str + ']/div/div/div[2]/div[1]/div[1]/div/div[2]/div[1]/a').text
        url = driver.find_element(By.XPATH, '//ol[@class="artdeco-list background-color-white _border-search-results_1igybl"]/li[' + pointer_str + ']/div/div/div[2]/div[1]/div[1]/div/div[2]/div[1]/a').get_attribute('href')

    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
    except TimeoutException:
        fullname = "TimeoutException"
        url = "TimeoutException"
        
    try: 
        #Wait until position appears in DOM.
        elem2 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="artdeco-list background-color-white _border-search-results_1igybl"]/li[' + pointer_str + ']/div/div/div[2]/div[1]/div[1]/div/div[2]/div[3]/span[2]')))
        #Get the position. 
        job_title = driver.find_element(By.XPATH, '//ol[@class="artdeco-list background-color-white _border-search-results_1igybl"]/li[' + pointer_str + ']/div/div/div[2]/div[1]/div[1]/div/div[2]/div[3]/span[2]').text

    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
    except TimeoutException:
        job_title = "TimeoutException"
        
    try:                               
        #Wait until company appears in DOM. 
        elem3 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="artdeco-list background-color-white _border-search-results_1igybl"]/li[' + pointer_str + ']/div/div/div[2]/div[1]/div[1]/div/div[2]/div[3]/a')))
        #Get the company. 
        company = driver.find_element(By.XPATH, '//ol[@class="artdeco-list background-color-white _border-search-results_1igybl"]/li[' + pointer_str + ']/div/div/div[2]/div[1]/div[1]/div/div[2]/div[3]/a').text
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
    except TimeoutException:
        company = "TimeoutException"

    try:       
        #Wait until location appears in DOM. 
        elem4 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="artdeco-list background-color-white _border-search-results_1igybl"]/li[' + pointer_str + ']/div/div/div[2]/div[1]/div[1]/div/div[2]/div[4]/span')))
        #Get the location. 
        location = driver.find_element(By.XPATH, '//ol[@class="artdeco-list background-color-white _border-search-results_1igybl"]/li[' + pointer_str + ']/div/div/div[2]/div[1]/div[1]/div/div[2]/div[4]/span').text
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
    except TimeoutException:
        location = "TimeoutException"

    person = Profile(company, fullname, job_title, location, url)
    leads.append(person)

    print(person._company + "%^&" + person._full_name + "%^&" + person._first_name + "%^&" + person._last_name + "%^&" + person._location + "%^&" + person._job_title + "%^&" + person._url)
    print("  ")


#This function opens a browser and logs into Linkedin.
#It navigates to Sales Navigator after logging in.    
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


#This function must be executed at the Sales Navigator home page.
#This function simply clicks the search bar and sends Enter to the browser.
#So, the browser moves to a page with all the filters for leads.
def start_empty_search_in_sales_nav(driver):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.ID,'global-typeahead-search-input')))
    
        search_bar = driver.find_element_by_id('global-typeahead-search-input');
        search_bar.send_keys(Keys.RETURN)

    except StaleElementReferenceException:
        driver.refresh()
        start_empty_search_in_sales_nav(driver)


def scroll_down(driver):

    #Select the div box containing the search results.    
    element = driver.find_element_by_xpath("/html/body/main/div/div[2]/div[2]")
    element.click()

    #Get the height of the div box containing the search results.
    height = driver.execute_script("return arguments[0].scrollHeight", element)

    #Divide the scroll height into 6 equal sections so driver can stop at each section.
    sect1 = height/6
    sect2 = sect1 * 2
    sect3 = sect1 * 3
    sect4 = sect1 * 4
    sect5 = sect1 * 5

    #Stop at each sections ending with the bottom of the page. Make sure all results load on the page.    
    driver.execute_script("arguments[0].scrollTo(0, " + str(sect1) + ");", element)
    time.sleep(1)
    driver.execute_script("arguments[0].scrollTo(0, " + str(sect2) + ");", element)
    time.sleep(1)
    driver.execute_script("arguments[0].scrollTo(0, " + str(sect3) + ");", element)
    time.sleep(1)
    driver.execute_script("arguments[0].scrollTo(0, " + str(sect4) + ");", element)
    time.sleep(1)
    driver.execute_script("arguments[0].scrollTo(0, " + str(sect5) + ");", element)
    time.sleep(1)
    driver.execute_script("arguments[0].scrollTo(0, " + str(height) + ");", element)
    time.sleep(1)

    time.sleep(1)
    
def get_num_of_search_result_pages(driver):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//main[@id="content-main"]/div/div[2]/div[2]/div/div[4]/div/ul')))
        
        page_num = driver.find_element(By.XPATH, '//main[@id="content-main"]/div/div[2]/div[2]/div/div[4]/div/ul/li[last()]/button').text
        
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
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//main[@id="content-main"]/div/div[2]/div[2]/div/ol')))
        html_list = driver.find_elements(By.XPATH, '//main[@id="content-main"]/div/div[2]/div[2]/div/ol/li')
        results_num = len(html_list)
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        get_num_of_search_results_in_current_page(driver)
        
    return results_num


def iterate_through_pages(driver, page_num):
    curr = 1

    #Start populating a list of all page urls, starting with current page url.
    page_urls.append(driver.current_url)

    #iterate_through_results(driver)
    while curr < page_num:
        curr+=1
        
        try:
            wait = WebDriverWait(driver, 10)
            element = wait.until(EC.presence_of_element_located((By.XPATH, '//main[@id="content-main"]/div/div[2]/div[2]/div/div[4]/div/ul[@class="artdeco-pagination__pages artdeco-pagination__pages--number"]')))
            nextPage = driver.find_element(By.XPATH, '//main[@id="content-main"]/div/div[2]/div[2]/div/div[4]/div/ul/li[@class="artdeco-pagination__indicator artdeco-pagination__indicator--number active selected ember-view"]/following-sibling::li/button')
            nextPage.send_keys(Keys.RETURN)
            
            time.sleep(2)
            page_urls.append(driver.current_url)
        except (StaleElementReferenceException , TimeoutException):
            curr-=1
            driver.refresh()


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

#Expand the filter if collapsed. If the filters are expanded, collapse the filters.
def expand_collapse_filters(driver):
    elem_xpath = '//main[@id="content-main"]/div/div[1]/div[1]/button'
    click_element_by_xpath(driver, elem_xpath)


def specify_a_seniority(driver, wanted_level):
    open_btn_xpath = '/html/body/main/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[3]/div/button'
    close_btn_xpath = '/html/body/main/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[3]/div[1]/button'
    results_xpath = '/html/body/main/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[3]/div[2]/div/ul'

    #Open the filter for seniority.
    click_element_by_xpath(driver, open_btn_xpath)
    time.sleep(1)
    
    seniority_list = driver.find_elements(By.XPATH, '/html/body/main/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[3]/div[2]/div/ul/li')

    for seniority in seniority_list:
        seniority_text = seniority.find_element(By.XPATH, './/div/span[1]').text
        if seniority_text == wanted_level:
            seniority_include = seniority.find_element(By.XPATH,'.//div/button[1]')
            seniority_include.click()
            break
            


def specify_a_company(driver, company):
    short_cut_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[1]/div/fieldset[1]' 
    open_btn_xpath =  '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[1]/div/fieldset[1]/div/button'
    close_btn_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[1]/div/fieldset[1]/div[1]/button'
    search_bar_xpath ='//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[1]/div/fieldset[1]/div[2]/div/div[1]/div/input'
    results_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[1]/div/fieldset[1]/div[2]/div/ul'
    
    #Open the filter for geography.
    click_element_by_xpath(driver, open_btn_xpath)

    #Type the location in the search bar in geography filter.
    type_in_search_bar_by_xpath(driver, search_bar_xpath, company)

    #Wait a second for the results to load from typing.
    time.sleep(1)

    #Get the number of results
    res_num = get_number_of_filter_search_results(driver, results_xpath)

    #If search returns only one result, li element does not need to specify [1] else, li need to be li[1] in xpath. 
    if int(res_num) > 1 :
        first_result_include_xpath = '/html/body/main/div/div[1]/form/div[1]/fieldset[1]/div/fieldset[1]/div[2]/div/ul/li[1]/div/button[1]'
        click_element_by_xpath(driver,first_result_include_xpath)
    elif int(res_num) == 1:
        first_result_include_xpath = '/html/body/main/div/div[1]/form/div[1]/fieldset[1]/div/fieldset[1]/div[2]/div/ul/li/div/button[1]'
        click_element_by_xpath(driver,first_result_include_xpath)
    else :
        print("No matching result for the function name or it can be a StaleElementReferenceException/TimeoutException.")

    click_element_by_xpath(driver,close_btn_xpath)


def specify_a_geography(driver, location):
    short_cut_xpath = '//main[@id="content-main"]/div/div[1]/form/div[2]/fieldset[1]/div/fieldset[3]' 
    open_btn_xpath = '//main[@id="content-main"]/div/div[1]/form/div[2]/fieldset[1]/div/fieldset[3]/div/button'
    close_btn_xpath = '//main[@id="content-main"]/div/div[1]/form/div[2]/fieldset[1]/div/fieldset[3]/div[1]/button'
    search_bar_xpath = '//main[@id="content-main"]/div/div[1]/form/div[2]/fieldset[1]/div/fieldset[3]/div[2]/div/div[1]/div/input'
    results_xpath = '//main[@id="content-main"]/div/div[1]/form/div[2]/fieldset[1]/div/fieldset[3]/div[2]/div/ul'

    #Open the filter for geography.
    click_element_by_xpath(driver, open_btn_xpath)

    #Type the location in the search bar in geography filter.
    type_in_search_bar_by_xpath(driver, search_bar_xpath, location)

    #Wait a second for the results to load from typing.
    time.sleep(1)

    #Get the number of results
    res_num = get_number_of_filter_search_results(driver, results_xpath)

    #If search returns only one result, li element does not need to specify [1] else, li need to be li[1] in xpath. 
    if int(res_num) > 1 :
        first_result_include_xpath = '/html/body/main/div/div[1]/form/div[2]/fieldset[1]/div/fieldset[3]/div[2]/div/ul/li[1]/div/button[1]'
        click_element_by_xpath(driver,first_result_include_xpath)
    elif int(res_num) == 1:
        first_result_include_xpath = '/html/body/main/div/div[1]/form/div[2]/fieldset[1]/div/fieldset[3]/div[2]/div/ul/li/div/button[1]'
        click_element_by_xpath(driver,first_result_include_xpath)
    else :
        print("No matching result for the function name or it can be a StaleElementReferenceException/TimeoutException.")

    click_element_by_xpath(driver,close_btn_xpath)


#Opens job title filter in Sales Navigator and it types in a job title.
#Then it select the topmost job title that appears from search. Clicks the function element.
#Then it closes the job title filter.
def specify_a_lead_job_title(driver, job_title):
    short_cut_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[2]' 
    open_btn_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[2]/div/button'
    close_btn_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[2]/div[1]/button'
    search_bar_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[2]/div[2]/div/div[1]/div/input'
    results_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[2]/div[2]/div/ul'

    #Open the filter for job title.
    click_element_by_xpath(driver, open_btn_xpath)
    
    #Type the job title in the search bar in job title filter.
    type_in_search_bar_by_xpath(driver, search_bar_xpath, job_title)

    #Wait a second for the results to load from typing.
    time.sleep(1)

    #Get the number of results
    res_num = get_number_of_filter_search_results(driver, results_xpath)

    #If search returns only one result, li element does not need to specify [1] else, li need to be li[1] in xpath. 
    if int(res_num) > 1 :
        first_result_include_xpath = '/html/body/main/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[2]/div[2]/div/ul/li[1]/div/button[1]'
        click_element_by_xpath(driver,first_result_include_xpath)
    elif int(res_num) == 1:
        first_result_include_xpath = '/html/body/main/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[2]/div[2]/div/ul/li/div/button[1]'
        click_element_by_xpath(driver,first_result_include_xpath)
    else :
        print("No matching result for the function name or it can be a StaleElementReferenceException/TimeoutException.")

    click_element_by_xpath(driver,close_btn_xpath)


#Opens function filter in Sales Navigator and it types in a function name.
#Then it select the topmost function that appears from search. Clicks the function element.
#Then it closes the function filter.
def specify_a_lead_function(driver, function_name):
    short_cut_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[1]'
    open_btn_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[1]/div/button'
    close_btn_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[1]/div[1]/button'
    search_bar_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[1]/div[2]/div/div[1]/div/input'
    results_xpath = '//main[@id="content-main"]/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[1]/div[2]/div/ul'

    click_element_by_xpath(driver, open_btn_xpath)

    #Type the function name in the search bar in function filter.
    type_in_search_bar_by_xpath(driver, search_bar_xpath, function_name)

    #Wait a second for the results to load from typing.
    time.sleep(1)

    #Get the number of results
    res_num = get_number_of_filter_search_results(driver, results_xpath)

    #If search returns only one result, li element does not need to specify [1] else, li need to be li[1] in xpath. 
    if int(res_num) > 1 :
        first_result_include_xpath = '/html/body/main/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[1]/div[2]/div/ul/li[1]/div/button[1]'
        click_element_by_xpath(driver,first_result_include_xpath)
    elif int(res_num) == 1:
        first_result_include_xpath = '/html/body/main/div/div[1]/form/div[1]/fieldset[2]/div/fieldset[1]/div[2]/div/ul/li/div/button[1]'
        click_element_by_xpath(driver,first_result_include_xpath)
    else :
        print("No matching result for the function name or it can be a StaleElementReferenceException/TimeoutException.")

    click_element_by_xpath(driver,close_btn_xpath)



def get_number_of_filter_search_results(driver, results_xpath):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, results_xpath)))
        results_list_element = driver.find_element(By.XPATH, results_xpath)
        
        return results_list_element.get_attribute('data-count')
    except(StaleElementReferenceException , TimeoutException) as e:
        print(e.message)
        return "0"

    return "0"


def type_in_search_bar_by_xpath(driver, search_bar_xpath, search_keywords):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, search_bar_xpath)))
        search_bar = driver.find_element(By.XPATH, search_bar_xpath)
        search_bar.send_keys(search_keywords)
    except(StaleElementReferenceException , TimeoutException) as e:
        print(e.message)
        


def click_element_by_xpath(driver, element_xpath):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, element_xpath)))
        btn = driver.find_element(By.XPATH, element_xpath)
        btn.click()

    except(StaleElementReferenceException , TimeoutException) as e:
        print(e.message)

    

# The following functions need to be rewritten to adapt to update UI of Sales Navigator
# def select_seniority_in_search(driver, level):
# def select_function_in_search(driver, category):
# def search_industry_in_search(driver, industry):
# def select_industry_in_search(driver):
# def open_search_results(driver, curr):
# def grab_details(driver):

main()
