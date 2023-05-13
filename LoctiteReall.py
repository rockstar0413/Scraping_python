# -*- coding: utf-8 -*-
"""
Created on Fri Sep  2 23:29:18 2022

@author: Admin
"""

from selenium import webdriver
from bs4  import BeautifulSoup
import pandas as pd
from selenium.webdriver import Chrome
import openpyxl
from pathlib import Path
import time
import requests
from time import sleep

driver = Chrome(executable_path='C:\\qqq.exe')

header = ['ITEM#', 'LOCTITE PART #', 'PRICE', 'PRODUCT NAME', 'PRODUCT DETAILS', 'UPC', 
          'STANDARD PACK', 'QTY AVAILABLE', 'Applicable Materials', 'Applications',
          'Color', 'IMAGE', 'Stock Status' ,'Duplicate', '', 'DESCRIPTION', 'Bullet PINTS', 'TDS', 'IMAGE1', 'IMAGE2' ]
totalData = []


xlsx_file = Path('', 'Loctite.xlsx')
wb_obj = openpyxl.load_workbook(xlsx_file) 
# Read the active sheet:
sheet = wb_obj.active
  
ITEM_names = []
parcelList = []

print("======================================================================================")


for i, row in enumerate(sheet.iter_rows(values_only=True)):
    
    ITEM = ""
    Loctite_Part = ""
    PartId = ""
    PRICE = ""
    Product_name = ""
    Product_detail = ""
    UPC = ""
    Unit_price = ""
    Standard_pack = ""
    QTY_avaliable = ""
    Category = ""
    Image = ""
    Applicable_Materials = ""
    Application = ""
    Color = ""
    Stock_Status = ""
    Duplicate = ""
    
    if i != 0:
        ITEM = row[0]
        Loctite_Part = row[1]
        PRICE = row[2]
        Product_name = row[3]
        Product_detail = row[4]
        UPC = row[5]
        Standard_pack = row[6]
        QTY_avaliable = row[7]
        Applicable_Materials = row[8]
        Application = row[9]
        Color = row[10]
        Image = row[11]
        Stock_Status = row[12]
        Duplicate = row[13]
        
        Description = ""
        Bullet_Points = ""
        TDS = ""
        Image1 = ""
        Image2 = ""
        
        print(Product_name)
        driver.get('https://www.henkel-adhesives.com/us/en/search.html?searchText='+Product_name)
        sleep(1)
        content = driver.page_source
        soup = BeautifulSoup(content,"lxml")

        searchResultList  =  soup.find('ul', attrs={'class': 'searchresults__boxlist'}).findAll('li', recursive=False)
        
        for searchItem in searchResultList:
            try:
                detailLink = searchItem.find('a', attrs={'class':'searchresults__categoryproductInfoItem'}).get('href')
                
                driver.get(detailLink)
                detailContent = driver.page_source
                detailSoup = BeautifulSoup(detailContent,"lxml")
                
                features = ""
                text = ""
                features = detailSoup.find('div', attrs={'class': 'product__features'}).text.strip()
                text = detailSoup.find('div', attrs={'class': 'product__descriptionText'}).text.strip()
                Description = features + text
                
                try:
                    bulletItem = detailSoup.find('div', attrs={'class': 'product__benefitsList'})
                    bulletList = bulletItem.findAll('li')
                    for bullet in bulletList:
                        Bullet_Points = Bullet_Points + ">" + bullet.text.strip() + "\n"
                except:
                    Bullet_Points = ""    
                    
                try: 
                    TDSItem = detailSoup.find('div', attrs={'class': 'product__featureButtons'})
                    TDS = TDSItem.find('a', attrs={'class', 'button__primaryOutline'} ).get('href')
                except:
                    TDS = ""
                    
                try:
                    ImageItem = detailSoup.find('div', attrs={'class': 'gallery__thumbnails'})
                    ImageList = ImageItem.findAll('img', attrs={'class': 'gallery__thumbnail-image'})
                    i = 0
                    for image in ImageList:
                        if i == 0:
                            Image1 = image.get('src').strip()
                        if i == 1:
                            Image2 = image.get('src').strip()
                        i = i + 1
                except:
                    Image1 = ""
                    Image2 = ""
                    
                print(Description)
                print(Bullet_Points)
                print(TDS)
                print(Image1)
                print(Image2)
        
                
                
                
                rowData = [ITEM, Loctite_Part,  PRICE, Product_name, Product_detail, UPC, Standard_pack.strip(), QTY_avaliable,
                           Applicable_Materials, Application, Color,  Image, Stock_Status, "", "", Description,
                           Bullet_Points, TDS, Image1, Image2]
                
                totalData.append(rowData)

               
                
                allDataFrame = pd.DataFrame()
                allDataFrame = pd.DataFrame(totalData, columns = header)
                filename = 'LoctiteReal.csv'
                allDataFrame.to_csv(filename,index=False )
                
                break
            except:
                continue
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        