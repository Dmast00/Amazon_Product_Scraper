import openpyxl
from openpyxl.cell import cell
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet import worksheet
from openpyxl.styles import alignment
from selenium import webdriver
from bs4 import BeautifulSoup
from openpyxl import workbook,load_workbook
from os import cpu_count, startfile,close
from selenium.common import exceptions
from selenium.webdriver.common import keys
from selenium.webdriver.firefox import options 
from selenium.webdriver.support import wait
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.keys import Keys
import pandas as pd 
import re
import time
import os
from time  import time,sleep
Menu_Options ={
    1:'Search From Excel',
    2:'Search Only one Product',
    3:'Exit'
}
url = "https://amazon.com"
def Print_Menu():
    for key in Menu_Options:
        print(key,'--',Menu_Options[key])

def Search_From_Excel():
    print('Searching from Excel')

def Count_items(driver ):
    items =driver.find_elements(By.CSS_SELECTOR,'.s-main-slot div.s-result-item.s-asin.sg-col-0-of-12.sg-col-16-of-20.sg-col')
    # Search for the Asin code and opens the product details page
    products_len =len(items) 
    return products_len

def Count_review_items(driver):
    # reviews = driver.find_elements(By.CSS_SELECTOR,'.a-section.review.aok-relative')
    # reviews = driver.find_elements(By.CSS_SELECTOR,'div.a-section.celwidget')
    reviews = driver.find_elements(By.CSS_SELECTOR,'div.a-section.review.aok-relative')
    reviews_id = []
    for rev in reviews:
        reviews_id.append(rev.get_attribute('id'))
    reviews_len = len(reviews)
    print(reviews_id)
    
    return reviews_id


def Select_Product_Name():
    cmd = 'cls'
    os.system(cmd)
    product_name = ''
    product_name = str(input('Enter the name of the prodcut: '))
    excel_name = product_name+'.xlsx'
    # driver = webdriver.Firefox()
    xl = openpyxl.Workbook(excel_name)
    xl.save(excel_name)
    # driver.get(url)
    product= ''
    driver = webdriver.Firefox()
    driver.get(url)

    search_bar = driver.find_element_by_id('twotabsearchtextbox')
    search_bar.send_keys(product_name)

    search_button = driver.find_element_by_xpath('//*[@id="nav-search-submit-button"]')
    search_button.click()
    count = 1
    Search_Only_One_Product(driver,product_name,count)
    return product_name

def Search_reviews(driver,wbFileName):
    workbook = Workbook()
    workbook = load_workbook(filename=wbFileName)
    sheet = workbook.active
    id_reviews = Count_review_items(driver)
    count = 1
    for id in id_reviews:
        rating_temp = driver.find_element_by_css_selector('#customer_review-{} > div:nth-child(2) > a:nth-child(1)'.format(id))
        rating = rating_temp.get_attribute('title')
        comment = driver.find_element_by_css_selector('#customer_review-{} > div:nth-child(5) > span:nth-child(1) > span:nth-child(1)'.format(id))
        sheet['D{}'.format(count)] = rating
        workbook.save(wbFileName)
        sheet['C']
        sheet['C{}'.format(count)] = comment
        workbook.save(wbFileName)
        count+=1

def Search_Only_One_Product(driver,product_name,count):

    Products_len = Count_items(driver)
    wbFileName = product_name+''+'.xlsx'
    workbook = Workbook()
    workbook = load_workbook(filename=wbFileName)
    print(wbFileName)
    sheet = workbook.active

    i = 1
    print(f'Count: {count}')
    while i <= Products_len:
        print(f'I value: {i}')
        sleep(2)
        try:
            search_asin = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[1]/div[1]/div/span[3]/div[2]/div[{}]'.format(i))
                                                        
        except:                                         
            print('ASIN not found')                                                                           
        asin =search_asin.get_attribute('data-asin')
        sleep(2)
        try:
            product = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[1]/div[1]/div/span[3]/div[2]/div[{}]/div/span/div/div/div[2]/div[2]/div/div/div[1]/h2/a/span'.format(i)).text
            sleep(1)                                                            
            print(product)
            sheet['A{}'.format(count)] = product
            workbook.save(wbFileName)
        except:
            pass
                                      
        if asin is not None:
            new_window = url+'/dp/'+asin
            driver.get(new_window)
            try:
                sheet['B{}'.format(count)] = new_window
                workbook.save(wbFileName)
                reviews = driver.find_element_by_css_selector('#reviews-medley-footer > div:nth-child(2) > a:nth-child(2)')                                             
                if reviews is None:
                    print('no comments section')                                         
                rev_url = reviews.get_attribute('href')
                driver.execute_script("window.open('{}');".format(rev_url))
                driver.switch_to.window(driver.window_handles[1])
                Search_reviews(driver,wbFileName)    
            except:
                try:
                    reviews = driver.find_element_by_css_selector('.a-link-emphasis')
                    if reviews is None:
                        print('no comments section') 
                    rev_url = reviews.get_attribute('href')
                    driver.execute_script("window.open('{}');".format(rev_url))
                    driver.switch_to_window(driver.window_handles[1])
                    Search_reviews(driver,wbFileName)
                except:
                    pass 
                

            print(i)
            driver.back()
        else:
            pass
        workbook.save(wbFileName)
        i +=1
        sleep(2)
        count +=1
    print(count)
    try:
        next_page = driver.find_element_by_css_selector('li.a-last > a:nth-child(1)')                                                        
        next_page.click()
        Search_Only_One_Product(driver,product_name,count)
        
    except:
        print('Not Found')
    try:
        next_page = driver.find_element_by_css_selector('a.s-pagination-item:nth-child(7)')                                                      
        next_page.click()
        Search_Only_One_Product(driver,product_name,count)
    except:
        print('Not found2')


def Closing_Script():
    exit()


if __name__ == '__main__':
    while (True):
        Print_Menu()
        option =''
        try:
            option = int(input('Enter an Option: '))
        except:
            print('Wrong input please enter a number!!')
        if option == 1:
                Search_From_Excel()
        elif option == 2:
                Select_Product_Name()
        elif option == 3:
                Closing_Script()
        else:
            print('Please enter a number between 1 and 3!!')



