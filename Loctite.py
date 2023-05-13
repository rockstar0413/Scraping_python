# -*- coding: utf-8 -*-
"""
Created on Fri Sep  2 23:29:18 2022

@author: Admin
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4  import BeautifulSoup
import pandas as pd
from selenium.webdriver import Chrome
import openpyxl
from pathlib import Path
import time
import requests
from time import sleep
import requests


driver = Chrome(executable_path='C:\\aaa.exe')

print("======================================================================================")

header = ['ITEM#', 'LOCTITE PART #', 'PRICE', 'PRODUCT NAME', 'PRODUCT DETAILS', 'UPC', 
          'STANDARD PACK', 'QTY AVAILABLE', 'Applicable Materials', 'Applications',
          'Color', 'Capacity Vol.[Nom]' , 'IMAGE', 'Stock Status' ,'Duplicate' ]
totalData = []
#driver = Chrome(executable_path='C:\\Users\\Admin\\.spyder-py3\\chromedriver.exe')

print("===========================================  log in ===========================")
driver.get("https://www.orsnasco.com/storefrontCommerce/home.do")
element = driver.find_element(By.ID, "usr_name")
element.send_keys("mpomeroy@strobelssupply.com")
element = driver.find_element(By.ID, "usr_password")
element.send_keys("Ors1234!")
element = driver.find_element(By.CLASS_NAME, "lgn_button")
element.click()

 
specificNameList = []

for page in range(32):

    driver.get('https://www.orsnasco.com/storefrontCommerce/breadcrumbSearch.do?displayThumbnail=true&currentPage='+str(page+1)+'&totalElements=797&elementsPerPage=25&pageClicked='+str(page+1)+'&breadcrumb_path=ORS_Catalog%2F%2F%2F%2FAttribSelect%3DBrand+%3D+%27LOCTITE%27')
   
    dataContent = driver.page_source
    dataSoup = BeautifulSoup(dataContent,"lxml")
    
    table  = dataSoup.find('table', attrs={'id': 'card-table'})
    
    tdList = table.findAll('td')
    
    for tdItem in tdList:
        
        ITEM = ""
        UPC = ""
        PartId = ""
        PRICE = ""
        Product_name = ""
        Product_detail = ""
        Unit_price = ""
        Standard_pack = ""
        QTY_avaliable = ""
        Category = ""
        Image = ""
        Applicable_Materials = ""
        Application = ""
        Color = ""
        Stock_Status = ""
        Capacity = ""

        if tdItem.get('name') != None and "itm_numlink" in tdItem.get('name').strip():
            
            detailLinkNumber = tdItem.get('name').strip().replace("itm_numlink","")
            detailPartId = tdItem.text.strip()
            driver.get('https://www.orsnasco.com/storefrontCommerce/itemDetail.do?item-id='+detailLinkNumber+'&item-number='+detailPartId)
          
            detailContent = driver.page_source
            detailSoup = BeautifulSoup(detailContent,"lxml")
            
          
            print('https://www.orsnasco.com/storefrontCommerce/itemDetail.do?item-id='+detailLinkNumber+'&item-number='+detailPartId)
            
            try:
                ITEM = detailSoup.find('td', attrs={'class': 'detail__item-no'}).text.strip()
                PartId = ITEM[4:len(ITEM)]
            except:
                ITEM = ""
            try:
                UPC = detailSoup.find('td', attrs={'class': 'detail__upc'}).text.strip()
            except:
                dnrlink = detailSoup.find('td', attrs={'class': 'dnrlink'})
                if dnrlink != None:

                    
                    try:
                        Product_name = detailSoup.find('td', attrs={'class': 'detail__desc'}).text.strip()
                    except:
                        Product_name = ""
                    
                    subLink = dnrlink.find('a').get('href')
                    driver.get('https://www.orsnasco.com'+subLink)
          
                    subDetailContent = driver.page_source
                    subDefetailSoup = BeautifulSoup(subDetailContent,"lxml")
                    DublicateITEM = ""
                    try:
                        DublicateITEM = subDefetailSoup.find('td', attrs={'class': 'detail__item-no'}).text.strip()
                        DublicatePartId = DublicateITEM[4:len(DublicateITEM)]
                    except:
                        DublicateITEM = ""
                    try:
                        stringLen = len(subDefetailSoup.find('td', attrs={'class': 'detail__pricesell'}).text.strip())
                    except:
                        stringLen = ""
                    try:
                        PRICE = subDefetailSoup.find('td', attrs={'class': 'detail__pricesell'}).text.strip()[1:stringLen-3]
                    except:
                        PRICE = ""
                    try:
                        stringLenUnit = len(subDefetailSoup.find('td', attrs={'class': 'detail__uprice'}).text.strip())
                    except:
                        stringLenUnit = ""
                    try:
                        Unit_price = subDefetailSoup.find('td', attrs={'class': 'detail__uprice'}).text.strip()[1:stringLenUnit-3]
                    except:
                        Unit_price = ""
                    try:
                        Standard_pack = subDefetailSoup.find('td', attrs={'class': 'details__stdpack'}).text.strip()
                    except:
                        Standard_pack = ""
                    try:
                        QTY_avaliable = subDefetailSoup.find('td', attrs={'class': 'text detail__avail'}).text.strip()
                    except:
                        QTY_avaliable = ""

                    
                    
                    table2 = subDefetailSoup.find('table', attrs={'class': "table2"})
                    try:
                        Category = table2.find('b').text
                    except:
                        Category = ""
                    try:
                        Product_detail_tag = table2.find('td')
                        for tag in ('b', 'br','b'):
                            try:
                                Product_detail_tag.find(tag).decompose()
                            except:
                                break
                        Product_detail = Product_detail_tag.text.strip()
                    except:
                        Product_detail = ""
                        
                    try:
                        ImageDiv = subDefetailSoup.find('div', attrs={'class': 'itemDetailImg'})
                        Image = "https://www.orsnasco.com/storefrontCommerce/"+ImageDiv.find('img').get('src')
                    except:
                        Image = ""
                        
                    specificList = subDefetailSoup.findAll('tr',  class_=['attrShade','attrNoShade'])
                   
                    specificValueList = [''] * len(specificNameList)
                    for specificItem in specificList:
                        
                        if specificItem.findAll('td')[0].text.strip() == "Applicable Materials":
                            Applicable_Materials = specificItem.findAll('td')[1].text.strip()
                        if specificItem.findAll('td')[0].text.strip() == "Applications":
                            Application = specificItem.findAll('td')[1].text.strip()
                        if specificItem.findAll('td')[0].text.strip() == "Stock Status":
                            Stock_Status = specificItem.findAll('td')[1].text.strip()
                        if specificItem.findAll('td')[0].text.strip() == "Color":
                            Color = specificItem.findAll('td')[1].text.strip()
                        if specificItem.findAll('td')[0].text.strip() == "Capacity Vol. [Nom]":
                            Capacity = specificItem.findAll('td')[1].text.strip()

                    print(ITEM)
                    print(UPC)
                    print(PartId)
                    print(PRICE)
                    print(Product_name)
                    print(Product_detail)
                    print(Unit_price)
                    print(Standard_pack)
                    print(QTY_avaliable)
                    print(Applicable_Materials)
                    print(Application)
                    print(Stock_Status)
                    print(Color)
                    print(Image)
            
                    
                    
                    rowData = [ITEM, PartId,  PRICE, Product_name, Product_detail, UPC, Standard_pack, QTY_avaliable,
                               Applicable_Materials, Application, Color, Capacity,  Image, Stock_Status, DublicateITEM]
                    
                    totalData.append(rowData)

              
                    allDataFrame = pd.DataFrame()
                    allDataFrame = pd.DataFrame(totalData, columns = header)
                    filename = 'Loctite.csv'
                    allDataFrame.to_csv(filename,index=False )
                    
                    continue

                    
                        
                UPC = ""
            try:
                stringLen = len(detailSoup.find('td', attrs={'class': 'detail__pricesell'}).text.strip())
            except:
                stringLen = ""
            try:
                PRICE = detailSoup.find('td', attrs={'class': 'detail__pricesell'}).text.strip()[1:stringLen-3]
            except:
                PRICE = ""
            try:
                stringLenUnit = len(detailSoup.find('td', attrs={'class': 'detail__uprice'}).text.strip())
            except:
                stringLenUnit = ""
            try:
                Unit_price = detailSoup.find('td', attrs={'class': 'detail__uprice'}).text.strip()[1:stringLenUnit-3]
            except:
                Unit_price = ""
            try:
                Standard_pack = detailSoup.find('td', attrs={'class': 'details__stdpack'}).text.strip()
            except:
                Standard_pack = ""
            try:
                QTY_avaliable = detailSoup.find('td', attrs={'class': 'text detail__avail'}).text.strip()
            except:
                QTY_avaliable = ""
            try:
                Product_name = detailSoup.find('td', attrs={'class': 'detail__desc'}).text.strip()
            except:
                Product_name = ""
            
            table2 = detailSoup.find('table', attrs={'class': "table2"})
            try:
                Category = table2.find('b').text
            except:
                Category = ""
            try:
                Product_detail_tag = table2.find('td')
                for tag in ('b', 'br','b'):
                    try:
                        Product_detail_tag.find(tag).decompose()
                    except:
                        break
                Product_detail = Product_detail_tag.text.strip()
            except:
                Product_detail = ""
                
            try:
                ImageDiv = detailSoup.find('div', attrs={'class': 'itemDetailImg'})
                Image = "https://www.orsnasco.com/storefrontCommerce/"+ImageDiv.find('img').get('src')
            except:
                Image = ""
                
            specificList = detailSoup.findAll('tr',  class_=['attrShade','attrNoShade'])
           
            specificValueList = [''] * len(specificNameList)
            for specificItem in specificList:
                
                if specificItem.findAll('td')[0].text.strip() == "Applicable Materials":
                    Applicable_Materials = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Applications":
                    Application = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Stock Status":
                    Stock_Status = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Color":
                    Color = specificItem.findAll('td')[1].text.strip()
               
                
            #========================================= henkel - adhesives. com   
            
            print(ITEM)
            print(UPC)
            print(PartId)
            print(PRICE)
            print(Product_name)
            print(Product_detail)
            print(Unit_price)
            print(Standard_pack)
            print(QTY_avaliable)
            print(Applicable_Materials)
            print(Application)
            print(Stock_Status)
            print(Color)
            print(Image)
    
            
            
            rowData = [ITEM, PartId,  PRICE, Product_name, Product_detail, UPC, Standard_pack, QTY_avaliable,
                       Applicable_Materials, Application, Color, Capacity, Image, Stock_Status, ""]
            
            totalData.append(rowData)

           
            
            allDataFrame = pd.DataFrame()
            allDataFrame = pd.DataFrame(totalData, columns = header)
            filename = 'Loctite.csv'
            allDataFrame.to_csv(filename,index=False )

