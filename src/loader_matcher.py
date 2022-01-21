#Using the linkedin url as the key, find a match, copy and paste contact details.
#I grabbed below from navigator.py and url_getter.py just in case I need it.
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
import xlrd
import pyperclip

def main():
    database_path = "./lead_database.xlsx"
    workbook_path = "./workbook_HCL_DX.xlsx"

    workbook = xlrd.open_workbook(workbook_path)
    print("Opened workbook!")
    database = xlrd.open_workbook(database_path)
    print("Opened database!")
    
    database_sheet = database.sheet_by_index(0)
    print("Created database table.")
    workbook_sheet = workbook.sheet_by_index(0)
    print("Created workbook table.")

    workbook.release_resources()
    database.release_resources()
    del database
    print("Database closed.")
    del workbook
    print("Workbook closed.")

    urls_list = []
    emails_list = []
    phones_list = []

    #Make three lists: urls, emails, phones from database excel file.
    for i in range(0, database_sheet.nrows):
        urls_list.append(database_sheet.row_values(i)[7])
        emails_list.append(database_sheet.row_values(i)[5]) 
        phones_list.append(database_sheet.row_values(i)[4])
    print("urls_list, emails_list, phones_list created.")

    #Just making sure the list is created correctly.
    #for i in range(0, database_sheet.nrows):
    #   print(str(urls_list[i]))
    
    #for i in range(0, database_sheet.nrows):
    #    print(str(emails_list[i]))
    
    #for i in range(0, database_sheet.nrows):
    #    print(str(phones_list[i]))
    

    #Use zip iterator to make two dictionaries: urls with emails, urls with phones.
    #Refer to https://www.kite.com/python/answers/how-to-create-a-dictionary-from-two-lists-in-python
    emails_zip_iterator = zip(urls_list, emails_list)
    phones_zip_iterator = zip(urls_list, phones_list)
    emails_dictionary = dict(emails_zip_iterator)
    print("Created email dictionary")
    phones_dictionary = dict(phones_zip_iterator)
    print("Created phone dictionary")

    #Make a list of urls from workbook which is the newly sourced leads from navigator.py
    #This list must only include the urls. Make sure to remove the heading like "Linkedin URL"
    #Refer to https://stackoverflow.com/questions/37403460/how-would-i-take-an-excel-file-and-convert-its-columns-into-lists-in-python
    workbook_urls_list = []
    for i in range(0, workbook_sheet.nrows):
        workbook_urls_list.append(workbook_sheet.row_values(i)[6])
    #Just making sure the list is created correctly.
    #for i in range(0, workbook_sheet.nrows):
    #   print(str(workbook_urls_list[i]))

    #Open workbook and create worksheet.
    #output file_name e.g "leads_.xlsx"
    output_filename = "output_matched_leads.xlsx"
    workbook = xlsxwriter.Workbook(output_filename)
    # sheet_name e.g "HCL Appscan 2022"
    worksheet = workbook.add_worksheet("emails_phones_urls")
    
    #Loop through urls in workbook_urls_list.   
    #Use dictionaries to check if dictionaries contain urls for each urls from the workbook.
    #If the dictionary contains the url, get the value(emails and phones).
    #Refer to https://www.w3schools.com/python/ref_dictionary_get.asp
    row = 1
    col = 0
    header = ["PHONE","EMAIL","URLS"]

    for hd in header:
        worksheet.write(0, col, hd)
        col+=1
    print("Wrote headers")
    
    for url in workbook_urls_list:
        #get method time complexity is O(1).
        #Refer to https://www.ics.uci.edu/~pattis/ICS-33/lectures/complexitypython.txt
        #Refer to https://wiki.python.org/moin/TimeComplexity
        email = emails_dictionary.get(url, "not found")
        phone = phones_dictionary.get(url, "not found")
        #Get the index of the url in the workbook.
        #Refer to https://www.programiz.com/python-programming/methods/list/index
        row_num = workbook_urls_list.index(url)
        #Put the phone number or/and email address into the column and row(index number) for email and phone.
        worksheet.write(row, 0, phone)
        worksheet.write(row, 1, email)
        worksheet.write(row, 2, url)
        print(phone + " %^& " + email + " %^& " + url )
        row+=1
    #Close the workbook and database.
        
    workbook.close()
    
main()
