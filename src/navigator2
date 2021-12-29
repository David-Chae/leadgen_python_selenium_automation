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
  
    driver.get(saved_search)
    
    #Zoom the browser to 60%.
    driver.execute_script("document.body.style.zoom='60%'")

    scroll_down(driver)

    time.sleep(5)

    #Close the browser and its process to prevent out of memory issue.
    driver.quit()

    print("---This program took %s seconds ---" % (time.time() - start_time))  


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

# The following functions need to be rewritten to adapt to update UI of Sales Navigator
# def temp_search(url):
# def select_seniority_in_search(driver, level):
# def select_function_in_search(driver, category):
# def search_industry_in_search(driver, industry):
# def select_industry_in_search(driver):
# def search_function_in_search(driver, function):
# def select_function_in_search(driver):
# def search_geography_in_search(driver, country):
# def select_geography_in_search(driver):
# def select_title_in_search(driver, title):
# def select_titles_in_search(driver, titles):
# def select_companies_in_search(driver, companies):
# def get_num_of_search_result_pages(driver):
# def get_num_of_search_results_in_current_page(driver):
# def iterate_through_pages(driver):
# def iterate_through_results(driver):
# def get_profile_data_from_search_result(driver, pointer):
# def open_search_results(driver, curr):
# def grab_details(driver):

main()
