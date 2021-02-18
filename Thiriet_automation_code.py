import threading
import time

from openpyxl import load_workbook
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np

#using selenium webdriver
from selenium import webdriver
from selenium.webdriver.common.keys import Keys 
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select

#URL
url_thiriet = 'https://livraison.thiriet.com/accueil/mes-services/mes-services,171,1149.html?&args=Y29tcF9pZD0xNjUmYWN0aW9uPWxpc3RlJmlkPSZjaGFuZ2VyX3ZpbGxlPTF8'

list_thiriet = 'brief/list-thiriet2.xlsx' #reading excel file
df_thiriet = pd.read_excel(list_thiriet)

postal_thiriet = df_thiriet['postal_code_thiriet']

city_thiriet = df_thiriet['commune_nom']
city_thiriet_new = []

for cit in city_thiriet:
    cit = cit.upper()
    cit = cit.replace('-',' ')
    print(cit)
    city_thiriet_new.append(cit)
    
    
def select_city(a):
    
    town = browser.find_element_by_id("ville_code_insee_iris") #selecting city
    town_drp = Select(town)
    name_town = a
    town_drp.select_by_visible_text(name_town)
    print(a)
    time.sleep(5)
    

def browsing_activity(val, num, city):
    
    nation = browser.find_element_by_id("pays_id")
    drp = Select(nation)
    drp.select_by_value("F") #selecting country
    
    post = browser.find_element_by_id("cp")
    
    post.clear()
    post.send_keys(str(num))
    post.send_keys(Keys.RETURN) #sending postal code

    print(str(num))
    
    time.sleep(5)
    
    thread = threading.Thread(target=select_city(city))
    
    thread.start()
    
    thread.join()
    
    submit_button = browser.find_element_by_xpath('/html/body/section/div/div[3]/div[2]/div[3]/form/div[2]/div/button')
    submit_button.click()
    
    message = browser.find_elements_by_xpath('/html/body/section/div/div[2]/div/div/p')[0].text
    
    browser.back()
    browser.back()
    
    time.sleep(5)
    
    post2 = browser.find_element_by_id("cp")
    post2.clear()
    
    time.sleep(5)
    
    return message
    
    



def main():
  
    count = 0
    response = []


    for p in postal_thiriet:

        try:

            town_now = city_thiriet_new[count]
            country = 'F'
            print(count)
            count = count + 1

            res = browsing_activity(country, p, town_now)
            response.append(res)

        except NoSuchElementException:

            try:
                browser.back()
                time.sleep(2)
                res = browsing_activity(country, p, town_now)
                response.append(res)

            except NoSuchElementException:
                pass


if __name__ == '__main__':
    browser = webdriver.Chrome() #launch browser
    browser.get(url_thiriet)
    main()