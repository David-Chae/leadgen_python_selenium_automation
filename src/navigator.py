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


class Lusha:
    
    def __init__(self, email, mobile):
        self._email = email
        self._mobile = mobile

    def get_email(self):
        return self._email
    
    def get_mobile(self):
        return self._mobile


def main():
    #try_lusha()
    test_search()
    
    

def test_search():
    #Take record of time that this program started running.
    start_time = time.time()

    driver = webdriver.Chrome()

    #Log into Linkedin Sales Navigator.
    loginToLinkedInSalesNav(driver)

    #Open an empty search page in Sales Navigator    
    startEmptySearchInSalesNav(driver)

    #Select Australia as geographical location of the leads.
    searchGeographyInSearch(driver, "Australia")

    #Select Chief Marketing Officer as a title in search.
    selectTitleInSearch(driver, 'Chief Marketing Officer')

    #Select Arts and Design as function in search.
    selectFunctionInSearch(driver, "Arts and Design")
    
    #Zoom the browser to 60%.
    driver.execute_script("document.body.style.zoom='60%'")

    #Following line of code has been commented out because the Linkedin returns too many request message.
    #selectCompaniesInSearch(driver, companies)

    #Get the number of pages in the search results.
    page_num = getNumOfSearchResultPages(driver)
    print("You are now ready to move on to working with " + str(page_num) + " pages.")

    #Following line of code has been commented out because test function has been refactored.
    #test(driver, "https://www.google.com.au")
    
    #Get the number of search results in the current page.
    results_num = getSearchResultsNumber(driver)
    print("You are now ready to move on to working with " + str(results_num) + " results in current page.")

    #Open each results in the search results and copy details into an object. Move to next page if necessary.
    iterateThroughPages(driver)
    print("All results have been printed.")

    #Write the copied details into an excel file.
    write_leads_to_excel_file("leads.xlsx", "Australia_CMO_Arts_N_Design")
    print("All leads data have been written to xlsx file.")
    
    time.sleep(6)
    driver.quit()

    print("---This program took %s seconds ---" % (time.time() - start_time))    


def try_lusha():
    
    #Load Lusha Extension to driver.
    extension='E:/python-selenium-code/10.3.2_0.crx'
    options = webdriver.ChromeOptions()
    options.add_extension(extension)
    driver = webdriver.Chrome(chrome_options=options)

    #Log into Linkedin Sales Navigator.
    loginToLinkedInSalesNav(driver)

    #Log into Lusha then closes the browser tab.
    loginToLusha(driver)

    #Wait until driver moves to far left first tab.
    time.sleep(2)
    #Refresh the tab because it takes refresh to update Lusha login.
    driver.refresh()
    #Give enough time for Lusha to appear on Linkedin page.
    time.sleep(2)

    #Open Lusha extension.
    open_lusha(driver)
    time.sleep(1)
    #Click to agree to Lusha privacy policy
    agree_to_lusha_privacy_policy(driver)
    time.sleep(1)

    #Get a Linkein account profile.
    driver.get("https://www.linkedin.com/in/john-m-b44ab3108/")
    time.sleep(2)
    
    #Open Lusha extension on thia profile.
    open_lusha(driver)
    time.sleep(1)

    try:
        if check_lusha_details_exist(driver):
            print("Lusha has contact details for this one.")
        else:
            print("Lusha has no contact details for this one.")
    except Exception as e:
        print(e)


#This function is still in progress, Still needs more code so it log into Lusha.
def loginToLusha(driver):

    driver.switch_to.window(driver.window_handles[1])
    driver.get("https://auth.lusha.com/login")

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@id="__next"]/div/div/div/span/form/span[1]/div/input')))

        login_form_id = driver.find_element(By.XPATH, '//div[@id="__next"]/div/div/div/span/form/span[1]/div/input')
        login_form_pw = driver.find_element(By.XPATH, '//div[@id="__next"]/div/div/div/span/form/span[2]/div/input')
        login_form_btn = driver.find_element(By.XPATH, '//div[@id="__next"]/div/div/div/span/form/button')

        file_id = open('lusha_id.txt','r')
        id = file_id.read()
        file_id.close()

        file_password = open('lusha_password.txt','r')
        password = file_password.read()
        file_password.close()

        login_form_id.send_keys(id)
        login_form_pw.send_keys(password)
        login_form_btn.send_keys(Keys.RETURN)
        time.sleep(15)
        
    except StaleElementReferenceException:
        driver.refresh()
        loginToLusha(driver)

    driver.close()
    driver.switch_to.window(driver.window_handles[0])


def loginToLinkedInSalesNav(driver):    
    
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
        loginToLinkedInSalesNav(driver)


def startEmptySearchInSalesNav(driver):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.ID,'global-typeahead-search-input')))
    
        search_bar = driver.find_element_by_id('global-typeahead-search-input');
        search_bar.send_keys(Keys.RETURN)

    except StaleElementReferenceException:
        driver.refresh()
        startEmptySearchInSalesNav(driver)
        

def selectFunctionInSearch(driver, category):

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
        selectFunctionInSearch(driver, category)


    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button')))
    
        function_country_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button')
        function_country_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        selectFunctionInSearch(driver, category)


def searchGeographyInSearch(driver, country):

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[5]')))
    
        geography = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]')
        geography.click()
        geography_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/input')
        geography_search_bar.send_keys(country)
        geography_search_bar.send_keys(Keys.RETURN)
        selectGeographyInSearch(driver)
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        searchGeographyInSearch(driver, country)


def selectGeographyInSearch(driver):    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/ol/li[1]/button')))
    
        geography_country_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/ol/li[1]/button')
        geography_country_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        selectGeographyInSearch(driver, country)


def selectTitleInSearch(driver, title):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[12]')))
    
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]//div[@class="ph4 pb4"]/input')
        filter_search_bar.send_keys(title)
        filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        

def selectTitlesInSearch(driver, titles):
    
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
        selectTitlesInSearch(driver, titles)


def selectCompaniesInSearch(driver, companies):
    
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
        selectCompaniesInSearch(driver, companies)

def getNumOfSearchResultPages(driver):
# I want to know the number of pages of search results
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]')))
    
        html_list = driver.find_element(By.XPATH, '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]')
        items = html_list.find_elements_by_tag_name("li")
        page_num = len(items)
        print("There are " + str(page_num) + " pages")
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        getNumOfSearchResultPages(driver)

    return page_num

def getSearchResultsNumber(driver):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]')))
        html_list = driver.find_elements(By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li')
        results_num = len(html_list)
        print("There are " + str(results_num) + " results in current page.")
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        getSearchResultsNumber(driver)
        
    return results_num
    

def iterateThroughPages(driver):
    
    page_num = getNumOfSearchResultPages(driver)
    curr = 1

    iterateThroughResults(driver)

    while curr < page_num:
        curr+=1

        try:
            wait = WebDriverWait(driver, 10)
            element = wait.until(EC.presence_of_element_located((By.XPATH, '//div/nav/ol[@class="search-results__pagination-list"]')))
        
            nextPage = driver.find_element(By.XPATH, '//div/nav/ol[@class="search-results__pagination-list"]/li['+ str(curr) + ']/button')
            nextPage.send_keys(Keys.RETURN)
            time.sleep(2)
            iterateThroughResults(driver)
        except (StaleElementReferenceException , TimeoutException):
            curr-=1


def iterateThroughResults(driver):
    results_num = getSearchResultsNumber(driver)
    scrollDown(driver)
    
    curr = 1
    while curr <= results_num:
        openSearchResults(driver, curr)
        curr+=1


def openSearchResults(driver, curr):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li['+ str(curr) + ']/div[2]/div/div/div/article/section[@class="result-lockup"]/div/div/dl/dt[@class="result-lockup__name"]/a')))

        url = driver.find_element(By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li['+ str(curr) + ']/div[2]/div/div/div/article/section[@class="result-lockup"]/div/div/dl/dt[@class="result-lockup__name"]/a').get_attribute('href')
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[1])
        driver.get(url);
        
        grabDetails(driver)

        time.sleep(2)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        openSearchResults(driver, curr)


def grabDetails(driver):

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
        grabDetails(driver)





def get_lusha():

    #Let's check if the lusha found any details. Did it return details?
    if check_lusha_could_not_find(driver):
        #If lusha
        return Lusha("no email", "no mobile")

    
    if check_lusha_show_btn(driver):
        #True means the button needs a click to load the email and phone number.
        btn = driver.find_element(By.XPATH, '//div[@class="contact-header"]/div/div[@class="contact-action-container"]/div[@class="save-to-action-buttons"]/button')
        btn.click()

    # I need to implement these functions.
    #Check if there are any details
    if check_lusha_details(driver):
        #Check if there are any email addresses.
        if check_lusha_email(driver):
            email = get_lusha_email(driver)
        #Check if there are any phone numbers.
        if check_lusha_mobile(driver):
            mobile = get_lush_mobile(driver)
            

def agree_to_lusha_privacy_policy(driver):

    try:
        lusha_frame = driver.find_element_by_id("LU__extension_iframe")
        driver.switch_to.frame(lusha_frame)
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//div[@class="privacy-approval-buttons-container"]/button')))
        button = driver.find_element(By.XPATH,'//div[@class="privacy-approval-buttons-container"]/button')
        ActionChains(driver).move_to_element(button).click(button).perform()
        driver.switch_to.default_content()
    except  Exception as e:
        print(e)

def open_lusha(driver):

    try:
        wait = WebDriverWait(driver, 10)
        elem_fullname = wait.until(EC.element_to_be_clickable((By.ID, 'LU__extension_badge_main')))
        button = driver.find_element(By.ID, 'LU__extension_badge_main')
        ActionChains(driver).move_to_element(button).click(button).perform()
        
    except(StaleElementReferenceException , TimeoutException):
        driver.refresh()
        open_lusha(driver)


def check_lusha_details_exist(driver):

    #Try and check if the lusha has save contact button.
    try:
        lusha_frame = driver.find_element_by_id("LU__extension_iframe")
        driver.switch_to.frame(lusha_frame)
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//div[@class="save-to-actions-buttons"]/button')))
        btn = driver.find_element(By.XPATH,'//div[@class="save-to-actions-buttons"]/button')
        driver.switch_to.default_content()
        return True        
    except NoSuchElementException:
        driver.switch_to.default_content()
        return False
    
    return False



def check_lusha_details(driver):
    try:
        contact_details = driver.find_elements(By.XPATH, '//div[@class="contact-header"]/ul[@class="contact-details"]/li')
        return True
    except NoSuchElementException:
        return False

def check_lusha_could_not_find(driver):
    try:
        empty_state = driver.find_element(By.XPATH, '//div[@class="enrich-empty-state-text"]/div[1]')
        return True
    except NoSuchElementException:
        return False


def check_lusha_show_btn(driver):

    try:
        show_contact_btn_text = driver.find_element(By.XPATH, '//div[@class="contact-header"]/div/div[@class="contact-action-container"]/div[@class="save-to-action-buttons"]/button/span').text
    except NoSuchElementException:
        return False

    if show_contact_btn_text == "Show contact":
        return True
    
    return False
    

def scrollDown(driver):
    
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






