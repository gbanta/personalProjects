#importing the necessary packages
import requests
from bs4 import BeautifulSoup
import pandas as pd
import sys
from datetime import datetime

#reading the excel file that has the list of zip codes 
excel = pd.read_excel(r"/Users/gbanta/Documents/repos/atbs/zips.xlsx")
zips = excel['Zip Codes']

#creating header object to get past wesbsite's captcha
headers = {
    'authority': 'www.niche.com',
    'cache-control': 'max-age=0',
    'sec-ch-ua': '"Google Chrome";v="87", " Not;A Brand";v="99", "Chromium";v="87"',
    'sec-ch-ua-mobile': '?0',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 11_0_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36',
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'sec-fetch-site': 'same-origin',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-user': '?1',
    'sec-fetch-dest': 'document',
    'accept-language': 'en-US,en;q=0.9',
}

#used to identify first zip code to create the final data frame
counter = 1

#iterating through the zip codes from the excel file
for zip in zips:

    print(zip)

    URL = 'http://www.niche.com/places-to-live/z/'+str(zip)+'/'

    page = requests.get(URL, headers=headers)                                       #get sites html

    soup = BeautifulSoup(page.content, 'html.parser')

    reportCard = soup.find(id='report-card')
    grades = reportCard.find_all('li', class_='ordered__list__bucket__item')        #obtaining grade html section

    about = soup.find(id='about')
    population = about.find('div', class_='scalar__value').text                     #obtaining population data

    about = soup.find(id='real-estate')
    realEstateValues = about.find_all('span')
    homeValue = realEstateValues[2].text
    rentValue = realEstateValues[4].text
    ownershipPercentages = about.find_all('div', class_='fact__table__row__value')
    rentPercent = ownershipPercentages[0].text
    ownPercent = ownershipPercentages[1].text                                       #obtaing real estate data

    residents = soup.find(id='residents')
    residentsData = residents.find_all('span')
    medianIncome = residentsData[2].text
    famWithChildPercent = residentsData[4].text                                     #obtaining residents data

    #create a list of the labels and a list of the data to be used when creating the data frame
    catList = ['Population', 'Median Home Value', 'Median Rent', 'Rent %', 'Own %', 'Median Income', '% Family with Children']
    report = [population, homeValue, rentValue, rentPercent, ownPercent, medianIncome, famWithChildPercent]

    #iterating through the report card to get the letter grades for each category
    for item in grades:

        category = item.find('div', class_='profile-grade__label')
        letter = category.findNext('div').contents[0]
        
        catList.append(category.text)
        report.append(letter)
    
    #if it is the first zip code, create the data frame, otherwise append the data frame
    if counter == 1:
        data = pd.DataFrame([report], columns=catList)
        counter = counter + 1
    else:
        data = data.append(pd.DataFrame([report], columns=catList))
    
# add the zip codes as the index of the list
data.index = (zips)

# print results to console and save to excel file 
print(data)
data.to_excel(r"/Users/gbanta/Documents/repos/atbs/ZipsOutput.xlsx"+datetime.now().strftime("%m-%d-%Y@%H.%M.%S")+".xls")