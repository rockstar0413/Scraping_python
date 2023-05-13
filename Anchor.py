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

header = ['ITEM#', 'UPC', 'PRICE', 'PRODUCT NAME', 'PRODUCT DETAILS', 'UNIT PRICE',
          'STANDARD PACK', 'QTY AVAILABLE', 'CATEGORY', 'IMAGE', '', ]
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

ITEM = ""
UPC = ""
PRICE = ""
Product_name = ""
Product_detail = ""
Unit_price = ""
Standard_pack = ""
QTY_avaliable = ""
Category = ""
Image = ""

#=======================
Abrasive_Material = ""
Arbor_Diam = ""
Brand = ""
Country_of_origin = ""
Dia = ""
Grit = ""
Hazmat = ""
MPI_Catalog = ""
Prop_65 = ""
Speed = ""
Stock_Status = ""
Tool_Shape = ""
detailLinkNumber= ""

#=======================

Anchor_Brand_Catalog_Page = ""
Arbor_Thread_TPI_Pitch = ""
Big_Catalog_Page = ""
MPI_Catalog_Page = ""
Packing_Type = ""
Trim_Length = ""
Type = ""
UNSPSC = ""
Wire_Material = ""
Wire_Size = ""

#======================
Block_Material = ""
Block_Width = ""
Bristle_Material = ""
Bristle_Rows = ""
Handle_Material = ""
Handle_Type = ""
Style = ""

#======================
Length = ""
Length_Per_Sheet = ""
Quantity = ""
Safety_Catalog_Page = ""
Width_per_Sheet = ""
Width = ""

#=======================
Applicable_Materials = ""
Application = ""
Body_Material = ""
Height = ""
Hub_Material = ""
Mounting = ""
Speed = ""
Used_With = ""

specificNameList = []

for page in range(58):
    
    driver.get('https://www.orsnasco.com/storefrontCommerce/breadcrumbSearch.do?displayThumbnail=true&currentPage='+str(page+1)+'&totalElements=1440&elementsPerPage=25&pageClicked='+str(page+1)+'&breadcrumb_path=ORS_Catalog%2F%2F%2F%2FAttribSelect%3DBrand+%3D+%27ANCHOR+BRAND%27')
    
    dataContent = driver.page_source
    dataSoup = BeautifulSoup(dataContent,"lxml")
    
    table  = dataSoup.find('table', attrs={'id': 'card-table'})
    
    tdList = table.findAll('td')
    
    for tdItem in tdList:
    
        if tdItem.get('name') != None and "item_desc_itemlink" in tdItem.get('name').strip():
            
            detailLinkNumber = tdItem.get('name').strip().replace("item_desc_itemlink","")
    
            driver.get('https://www.orsnasco.com/storefrontCommerce/itemDetail.do?item-id='+detailLinkNumber)
    
            detailContent = driver.page_source
            detailSoup = BeautifulSoup(detailContent,"lxml")
            try:
                ITEM = detailSoup.find('td', attrs={'class': 'detail__item-no'}).text.strip()
            except:
                ITEM = ""
            try:
                UPC = detailSoup.find('td', attrs={'class': 'detail__upc'}).text.strip()
            except:
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
            Product_name = detailSoup.find('td', attrs={'class': 'detail__desc'}).text.strip()
            
            
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
                
    
            ImageDiv = detailSoup.find('div', attrs={'class': 'itemDetailImg'})
            Image = "https://www.orsnasco.com/storefrontCommerce/"+ImageDiv.find('img').get('src')
            
            specificList = detailSoup.findAll('tr',  class_=['attrShade','attrNoShade'])
           
            specificValueList = [''] * len(specificNameList)
            for specificItem in specificList:
                
                if  specificItem.findAll('td')[0].text.strip() in specificNameList:
                    specificValueList[specificNameList.index(specificItem.findAll('td')[0].text.strip())] = specificItem.findAll('td')[1].text.strip()
                else:
                    specificNameList.append(specificItem.findAll('td')[0].text.strip())
                    specificValueList.append(specificItem.findAll('td')[1].text.strip())
                
                
                if specificItem.findAll('td')[0].text.strip() == "Anchor Brand Catalog Page#":
                    Anchor_Brand_Catalog_Page = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Big Catalog Page#":
                    Big_Catalog_Page = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Block Material":
                    Block_Material = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Block Width [Nom]":
                    Block_Width = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Brand":
                    Brand = specificItem.findAll('td')[1].text.strip();
                if specificItem.findAll('td')[0].text.strip() == "Bristle Material":
                    Bristle_Material = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Bristle Rows":
                    Bristle_Rows = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Country of Origin":
                    Country_of_origin = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Handle Material":
                    Handle_Material = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Handle Type":
                    Handle_Type = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Hazmat":
                    Hazmat = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "MPI Catalog":
                    MPI_Catalog = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "MPI Catalog Page#":
                    MPI_Catalog_Page = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Prop 65":
                    Prop_65 = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Stock Status":
                    Stock_Status = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Style":
                    Style = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Type":
                    Type = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "UNSPSC":
                    UNSPSC = specificItem.findAll('td')[1].text.strip()
                    
                if specificItem.findAll('td')[0].text.strip() == "Arbor Thread - TPI or Pitch":
                    Arbor_Thread_TPI_Pitch = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Dia. [Nom]":
                    Dia = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Packing Type":
                    Packing_Type = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Trim Length [Nom]":
                    Trim_Length = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Wire Material":
                    Wire_Material = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Wire Size [Nom]":
                    Wire_Size = specificItem.findAll('td')[1].text.strip()
                    
                if specificItem.findAll('td')[0].text.strip() == "Length [Nom]":
                    Length = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Length per Sheet":
                    Length_Per_Sheet = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Quantity":
                    Quantity = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Safety Catalog Page#":
                    Safety_Catalog_Page = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Width Per Sheet":
                    Width_per_Sheet = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Width [Nom]":
                    Width = specificItem.findAll('td')[1].text.strip()
                    
                if specificItem.findAll('td')[0].text.strip() == "Applicable Materials":
                    Applicable_Materials = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Applications":
                    Application = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Body Material":
                    Body_Material = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Height [Nom]":
                    Height = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Hub Material":
                    Hub_Material = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Mounting":
                    Mounting = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Speed [Max]":
                    Speed = specificItem.findAll('td')[1].text.strip()
                if specificItem.findAll('td')[0].text.strip() == "Used With":
                    Used_With = specificItem.findAll('td')[1].text.strip()
                    
                    
                
            print(ITEM)
            print(UPC)
            print(PRICE)
            print(Product_name)
            print(Product_detail)
            print(Unit_price)
            print(Standard_pack)
            print(QTY_avaliable)
            print(Category)
            print(Image)
    
            print("==================")
            print(specificNameList)
            print(specificValueList)
            
            header = ['ITEM#', 'UPC', 'PRICE', 'PRODUCT NAME', 'PRODUCT DETAILS', 'UNIT PRICE',
                      'STANDARD PACK', 'QTY AVAILABLE', 'CATEGORY', 'IMAGE', '', ]
            
            rowData = [ITEM, UPC, PRICE, Product_name, Product_detail, Unit_price, Standard_pack, QTY_avaliable,
                       Category, Image, ""]
            
            for nameItem in specificNameList:
                header.append(nameItem)
            for valueItem in specificValueList:
                rowData.append(valueItem)
                
            totalData.append(rowData)
            
            realData = []
            for totalItem in totalData:
                if len(totalItem) < len(header):
                    tempItem = totalItem
                    for x in range (1, len(header) - len(totalItem)):
                        tempItem.append("")
                    realData.append(tempItem)
                else:
                    realData.append(totalItem)
           
            try:
                allDataFrame = pd.DataFrame()
                allDataFrame = pd.DataFrame(realData, columns = header)
                filename = 'Anchor.csv'
                allDataFrame.to_csv(filename,index=False )
            except:
                print("")
