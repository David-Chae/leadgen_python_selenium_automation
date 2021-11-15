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
    #selectCompaniesInSearch(driver, companies)
    page_num = getSearchResultPageNumber(driver)
    print("You are now ready to move on to working with " + str(page_num) + " pages.")
    #test(driver, "https://www.google.com.au")
    results_num = getSearchResultsNumber(driver)
    print("You are now ready to move on to working with " + str(results_num) + " results in current page.")

    height = driver.execute_script("return document.documentElement.scrollHeight")
    firstQuarter = height/4
    halfway = height/2
    lastQuarter = firstQuarter + halfway
    
    driver.execute_script("window.scrollTo(0, " + str(firstQuarter) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(halfway) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(lastQuarter) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    
    curr = 1
    while curr <= results_num:
        openSearchResults(driver, curr)
        curr+=1

    
    time.sleep(6)
    driver.quit()


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
    except:
        driver.quit()
    finally:
        search_bar = driver.find_element_by_id('global-typeahead-search-input');
        search_bar.send_keys(Keys.RETURN)


def selectGeographyInSearch(driver, country):

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[5]')))
    except:
        driver.quit()
    finally:
        geography = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]')
        geography.click()

        geography_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/input')
        geography_search_bar.send_keys(country)
        geography_search_bar.send_keys(Keys.RETURN)


    try:
        wait = WebDriverWait(driver, 1)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/ol/li[1]/button')))
    except:
        driver.quit()
    finally:
        geography_country_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/ol/li[1]/button')
        geography_country_btn.send_keys(Keys.RETURN)


def selectTitleInSearch(driver, title):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[12]')))
    except:
        driver.quit()
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
    except:
        driver.quit()
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
    except:
        driver.quit()
    finally:
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]//div[@class="ph4 pb4"]/input')
        for company in companies:
            filter_search_bar.send_keys(company)
            filter_search_bar.send_keys(Keys.RETURN)

def getSearchResultPageNumber(driver):
# I want to know the number of pages of search results
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]')))
    except:
        driver.quit()
    finally:
        html_list = driver.find_element(By.XPATH, '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]')
        items = html_list.find_elements_by_tag_name("li")
        page_num = len(items)
        print("There are " + str(page_num) + " pages")

    return page_num

def getSearchResultsNumber(driver):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]')))
    except:
        driver.quit()
    finally:
        html_list = driver.find_elements(By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li')
        results_num = len(html_list)
        print("There are " + str(results_num) + " results in current page.")

    return results_num

def openSearchResults(driver, curr):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li['+ str(curr) + ']/div[2]/div/div/div/article/section[@class="result-lockup"]/div/div/dl/dt[@class="result-lockup__name"]/a')))
    except:
        driver.quit()
    finally:
        url = driver.find_element(By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li['+ str(curr) + ']/div[2]/div/div/div/article/section[@class="result-lockup"]/div/div/dl/dt[@class="result-lockup__name"]/a').get_attribute('href')
        test(driver, url)

def grabDetails(driver):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dt/span')))
    except:
        driver.quit()
    finally:
        fullname = driver.find_element(By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dt/span').text
        position = driver.find_element(By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dt').text
        location = driver.find_element(By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dd[@class="mt4 mb0"]/div').text
        print("Fullname : " + fullname)
        print("Position : " + position)
        print("Location : " + location)
                             
    

def test(driver, url):
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    driver.get(url);
    grabDetails(driver)
    time.sleep(3)
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    


main()
