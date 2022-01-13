#Generate keys(=linkedin url) for each sales navigator url.
#So, then I can use these keys to automate the database.
#Automation include
# 1. updating the lead's company and job title.
# 2. loading existing leads details with url.

#I grabbed below from navigator.py just in case I need it.
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException
from nameparser import HumanName
import xlsxwriter
import time
import re, string
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException

# Necessary library.
# Must install pandas with command pip install pandas xlrd openpyxl
import pandas as pd
import xlrd
import openpyxl
import pyperclip

sales_list = []
linkedin_list = []
temp_list = []

def test(driver):
    name = "Christopher Neal "
    job_title = "Global Chief Information Security Officer"
    search_keywords = name + job_title
    
    #Get variables to use beforehand.
    search_bar_xpath = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div[2]/div[2]/input'
    wait = WebDriverWait(driver, 10)
    
    #Bring google home page.
    driver.get("https://www.google.com/")
    elem = wait.until(EC.presence_of_element_located((By.XPATH, search_bar_xpath)))
    search_bar = driver.find_element(By.XPATH, search_bar_xpath)
    stale_element = False

    try:
        #Type in the keywords and send enter to search.
        search_bar.send_keys("linkedin " + search_keywords)
        search_bar.send_keys(Keys.RETURN)
        time.sleep(1)
    except StaleElementReferenceException:
        #This basically handles nothing and moves to next try statement.
        #This variable might come in handy later on.
        stale_element = True

    time.sleep(60)

    try:
        #Select the first google search result and get its url. Then print it.
        first_result_h3 = driver.find_element_by_tag_name('h3')
        first_result = first_result_h3.find_element(By.XPATH, "./parent::a")
        #find_element(By.XPATH, first_result_xpath)
        url = first_result.get_attribute('href')
    except NoSuchElementException:
        url = "not found"
    except StaleElementReferenceException:
        url = "stale element referenced."
        
    print(url)
    time.sleep(3)
    
def get_linkedin_url_through_google(driver, search_keywords):
    #Get variables to use beforehand.
    #search_bar_xpath = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div[2]/div[2]/input'
    search_bar_xpath = '/html/body/div[4]/div[2]/form/div[1]/div[1]/div[2]/div[2]/div/div[2]/input'
    wait = WebDriverWait(driver, 10)
    
    #Bring google home page.
    #driver.get("https://www.google.com/")
    elem = wait.until(EC.presence_of_element_located((By.XPATH, search_bar_xpath)))
    search_bar = driver.find_element(By.XPATH, search_bar_xpath)
    stale_element = False

    try:
        #Type in the keywords and send enter to search.
        search_bar.send_keys(Keys.CONTROL, 'a')
        search_bar.send_keys(Keys.DELETE)
        search_bar.send_keys("linkedin " + search_keywords)
        search_bar.send_keys(Keys.RETURN)
        time.sleep(1)
    except StaleElementReferenceException:
        #This basically handles nothing and moves to next try statement.
        #This variable might come in handy later on.
        stale_element = True

    try:
        #Select the first google search result and get its url. Then print it.
        first_result_h3 = driver.find_element_by_tag_name('h3')
        first_result = first_result_h3.find_element(By.XPATH, "./parent::a")
        #find_element(By.XPATH, first_result_xpath)
        url = first_result.get_attribute('href')
    except NoSuchElementException:
        url = "not found"
    except StaleElementReferenceException:
        url = "stale element referenced."
        
    print(url)
    
def main():
    #Take record of time that this program started running.
    start_time = time.time()

    #This is the location of the excel file.
    location = "urls_to_convert2.xlsx"
    print("Location of file is " + location)

    #Open the excel file.
    wb = xlrd.open_workbook(location)
    print("Workbook opened.")

    #Access first sheet and set it as sheet.
    sheet = wb.sheet_by_index(3)
    print("Third sheet accessed.")

    for i in range(sheet.nrows):
        first_name = sheet.cell_value(i, 1)
        last_name = sheet.cell_value(i, 2)
        job_title = sheet.cell_value(i, 3)
        search_keywords = first_name + " " + last_name + " " + job_title
        temp_list.append(search_keywords)
        print(search_keywords)

    print(str(len(temp_list)) + " urls will be printed.")

    driver = webdriver.Chrome()

    test(driver)
    
    i = 0
    while i < len(temp_list):
        #get_linkedin_url_from_linkedin_url(driver, temp_list[i])
        get_linkedin_url_through_google(driver, temp_list[i])
        i+=1

    driver.quit()

    print("---This program took %s seconds ---" % (time.time() - start_time)) 




#This function contains code I wrote in main function to get linkedin urls from sales navigator urls.
def trial_sales_nav_to_linkedin_url():
    #This is the location of the excel file.
    location = "IT_NSW_CXO.xlsx"
    print("Location of file is " + location)

    #Open the excel file.
    wb = xlrd.open_workbook(location)
    print("Workbook opened.")

    #Access first sheet and set it as sheet.
    sheet = wb.sheet_by_index(0)
    print("First sheet accessed.")

    for i in range(sheet.nrows):
        val = sheet.cell_value(i, 6)
        sales_list.append(val)
        print(val)

    sales_list.remove('LINKEDIN URL')

    driver = webdriver.Chrome()
    #Log into Linkedin Sales Navigator.
    log_into_linkedin(driver)
    
    counter = 0
    for sales_url in sales_list:
        counter+= 1
        if counter > 50:
            driver.quit()
            driver = webdriver.Chrome()
            #Log into Linkedin Sales Navigator.
            log_into_linkedin(driver)
            counter = 1
            
        get_linkedin_url_from_sales_nav(driver, sales_url)

    write_leads_to_excel_file("linkedin_urls", location)     


def get_linkedin_url_from_sales_nav(driver, sales_url):
    wait = WebDriverWait(driver, 10)
    timeout = False
    
    try:     
        driver.get(sales_url)
        li_menu_btn_xpath = '/html/body/main/div[1]/div[1]/div/div[2]/div[1]/div[3]/button/li-icon'
        elem = wait.until(EC.presence_of_element_located((By.XPATH, li_menu_btn_xpath)))
        li_menu_btn = driver.find_element(By.XPATH, li_menu_btn_xpath)
        li_menu_btn.click()

        copy_linkedin_url_btn_xpath = '/html/body/main/div[1]/div[1]/div/div[2]/div[1]/div[3]/div/div[1]/div/ul/li[4]/div'
        elem = wait.until(EC.presence_of_element_located((By.XPATH, copy_linkedin_url_btn_xpath)))
        copy_linkedin_url_btn = driver.find_element(By.XPATH, copy_linkedin_url_btn_xpath)

        copy_linkedin_url_btn.click()
    except ElementClickInterceptedException:
        get_linkedin_url_from_sales_nav(driver, sales_url)
    except TimeoutException:
        timeout = True
                
    if timeout == False:
        url = pyperclip.paste()
        print(url)
        linkedin_list.append(url)
    else:
        url = "Timeout:" + sales_url
        print(url)
        linkedin_list.append(url)


def get_linkedin_url_from_linkedin_url(driver, url):
    wait = WebDriverWait(driver, 10)
    driver.get(url)
    print(driver.current_url)
    

#This function opens a browser and logs into Linkedin.
#It navigates to Sales Navigator after logging in.    
def log_into_linkedin(driver):
    
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
        
    except StaleElementReferenceException:
        driver.refresh()
        log_into_linked_in_sales_nav(driver)

def write_leads_to_excel_file(file_name, sheet_name):
    # file_name e.g "leads_.xlsx"
    workbook = xlsxwriter.Workbook(file_name)
    # sheet_name e.g "HCL Appscan 2022"
    worksheet = workbook.add_worksheet(sheet_name)

    row = 1
    col = 0

    header = ["LINKEDIN URL"]

    for hd in header:
        worksheet.write(0, col, hd)
        col+=1

    for url in linkedin_list:
        worksheet.write(row, 0, url)
        print(str(row) + " url(s) has been written to file.")
        row+=1        
    
    workbook.close()

main()
#test()
