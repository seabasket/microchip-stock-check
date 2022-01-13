"""
Edited by Sebastian Rowe. 

 * Added Arrow, Avnet Americas, Newark, and Future Electronics support.
 * Added menu  that allows for searching through different MPN search sets.
 * Added ctrl-c instant quit
 
"""

import requests
import os
import sys
from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd
import re

try:
    #User Agent String - makes webpage think script is a browser
    header = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.75 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest"
    }

    startswith = 'MPN' #all search sets should start with 'startswith' value.


    path = os.path.realpath(__file__)
    scriptname = '/'  + str(os.path.basename(__file__))
    path = path.replace(scriptname, "")

    searchsets = []

    print('press ctrl-c at any time to terminate.')
    print(path)
    print(os.getcwd())
    print(os.listdir(path))
    print(os.listdir(os.getcwd()))

    for x in os.listdir(path):                                              #populates searchsets with all searchsets in directory
        if x.endswith('.xlsx') and x.startswith(startswith):                #must be an excel file and start with 'startswith'
            searchsets.append(x)

    if not searchsets:                                                      #if there are no searchsets in the directory
        print('[NO SEARCH SETS] - PLEASE ADD A SEARCH SET TO THE DIRECTORY')
        sys.exit()                                                          #immediately ends program

    while(True):                                                            #main menu
        print('Please select your search set.')                                 
        i = 0
        for x in searchsets:
            i += 1
            print(f"[{i}] - {str(x)}")
        index = int(input('\nPlease input your desired search set --> '))
        if index >  0 and index <= len(searchsets):
            searchset = str(searchsets[index - 1])
            print('searchset = ' + searchset)
            break
        else:
            print(f'[INVALID INPUT] - PLEASE ENTER A VALUE BETWEEN 1 and {len(searchsets)}.')
            continue



    df = pd.read_excel(searchset)                                             #Read in PolarFire MPNs

    df['Microchip Stock:'] = ""                                                 #Create blank column to store stock
    df['Digikey Stock:'] = ""                                                   #Create blank column to store stock
    df['Mouser Stock:'] = ""                                                    #Create blank column to store stock
    df['Arrow Stock:'] = ""                                                     #Create blank column to store stock
    df['Avnet Americas Stock:'] = ""                                            #Create blank column to store stock
    df['Newark Stock:'] = ""                                                    #Create blank column to store stock
    ## df['Verical Stock:'] = ""                                                   #Create blank column to store stock
    df['Future Electronics Stock:'] = ""                                        #Create blank column to store stock

    ndevices = len(df.index)
    for i in range(ndevices):                                                   #Iterate through list of devices on findchips.com
        partnum = df.loc[i, 'Device']
        url = "https://www.findchips.com/search/" + partnum                     #Open device page on findchips.com
        print(str(i + 1) + " --> " + str(url))                                  #Print loop iteration and URL being scraped
        response = requests.get(url, headers=header)                            #Load webpage
        soup = BeautifulSoup(response.text,'lxml')                              #Parse webpage
        mchpdiv = soup.find("div", {"id": "list1612"})                          #Find Microchip Direct section
        dkdiv = soup.find("div", {"id": "list1588"})                            #Find Digikey section
        mouserdiv = soup.find("div", {"id": "list1577"})                        #Find Mouser section
        arrowdiv = soup.find("div", {"id": "list1538"})                         #Find Arrow section
        avnetusdiv = soup.find("div", {"id": "list313766971"})                  #Find Avnet section
        newarkdiv = soup.find("div", {"id": "list1561"})                        #Find Newark section
        vericaldiv = soup.find("div", {"id": "list2167609"})                    #Find Verical section
        fediv = soup.find("div", {"id": "list1555"})                            #Find Future Electronic section
        

        #Microchip Direct Checker
        if(mchpdiv):                                                            #If Microchip Direct stocks this part
            stock = mchpdiv.find('td',{"class":"td-stock"}).getText().strip()   #Get Stock
            if stock == "Temporarily Out of Stock – Alternates Available":      #Set out of stock to 0
                stock = 0
            else:                                                               #Strip substring from stock number
                stock = int(re.sub('[^0-9]','', stock))                         #Removes all non-integers from 'stock', is an int.
            df.loc[i,'Microchip Stock:'] = stock                                #Add stock level to dataframe
        else:
            df.loc[i,'Microchip Stock:'] = 0                                    #Set stock to Zero if part is not found.
        
        #Digikey Checker    
        if(dkdiv):                                                              #If Digikey stocks this part
            stock = dkdiv.find('td',{"class":"td-stock"}).getText().strip()     #Get Stock
            if stock == "On Order" or stock == "Temporarily Out of Stock" or stock == "Limited Supply - Call":      #Set out of stock to 0
                stock = 0
            else:                                                               #Strip substring from stock number
                stock = int(re.sub('[^0-9]','', stock))                         #Removes all non-integers from 'stock', is an int.
            df.loc[i,'Digikey Stock:'] = stock
        else:
            df.loc[i,'Digikey Stock:'] = 0                                      #Set stock to Zero if part is not found.

        #Mouser Checker    
        if(mouserdiv):                                                          #If Mouser stocks this part
            stock = mouserdiv.find('td',{"class":"td-stock"}).getText().strip() #Get Stock    
            stock = int(stock)
            df.loc[i,'Mouser Stock:'] = stock                                   #Add stock level to dataframe
        else:
            df.loc[i,'Mouser Stock:'] = 0                                       #Set stock to Zero if part is not found.

        #Arrow Checker
        if(arrowdiv):
            partattrstr = "[data-mfrpartnumber=" + partnum + "]"
            specificpartnum = arrowdiv.select(partattrstr)  
            if not specificpartnum: #If specificpartnum is empty
                df.loc[i,'Arrow Stock:'] = 0
            else:
                for market in specificpartnum: #loop thru rows with the specific part number ewanted
                    stock = market.find('td',{"class":"td-stock"}).getText().strip()
                    if 'Americas' not in stock: #ensure that the market is not outside of the americas
                        stock = 0
                        continue
                    else:
                        stock = int(re.sub('[^0-9]','', stock))                     #Removes all non-integers from 'stock', is an int.
                        break
                df.loc[i,'Arrow Stock:'] = stock
        else:
            df.loc[i,'Arrow Stock:'] = 0

        #Avnet Americas Checker    
        if(avnetusdiv):                                                         #If Avnet Americas stocks this part
            stock = avnetusdiv.find('td',{"class":"td-stock"}).getText().strip()    #Get Stock    
            stock = int(stock)
            df.loc[i,'Avnet Americas Stock:'] = stock                           #Add stock level to dataframe
        else:
            df.loc[i,'Avnet Americas Stock:'] = 0                               #Set stock to Zero if part is not found.

        #Newark Checker    
        if(newarkdiv):                                                          #If Newark stocks this part
            stock = newarkdiv.find('td',{"class":"td-stock"}).getText().strip() #Get Stock    
            stock = int(stock)
            df.loc[i,'Newark Stock:'] = stock                                   #Add stock level to dataframe
        else:
            df.loc[i,'Newark Stock:'] = 0                                       #Set stock to Zero if part is not found.
            
        #ADD VERICAL HERE
            
        #Future Electronics Checker    
        if(fediv):                                                              #If Future Electronics stocks this part
            stock = fediv.find('td',{"class":"td-stock"}).getText().strip()     #Get Stock
            stock = int(re.sub('[^0-9]','', stock))                             #Removes all non-integers from 'stock', is an int.
            df.loc[i,'Future Electronics Stock:'] = stock                       #Add stock level to dataframe
        else:
            df.loc[i,'Future Electronics Stock:'] = 0                           #Set stock to Zero if part is not found.
                
            
    df["Total"] = df.drop(['Device', 'Package'],axis=1).sum(axis=1,numeric_only=False)                      #Sum totals

    df.to_excel('APA-Channel-Stock-{}.xlsx'.format(datetime.today().strftime('%Y%m%d-%H%M%S')),index=False) #Created Excel file

    print("Done!")   

except KeyboardInterrupt:
    pass 
