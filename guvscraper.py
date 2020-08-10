import sys
import requests
from bs4 import BeautifulSoup
import pandas as pd
from requests.api import head
import xlsxwriter
from datetime import datetime


company = "volkswagen"
base_url = "https://www.finanzen.net/bilanz_guv/"
user_agent = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36"}


if len(sys.argv) <= 2:
    if len(sys.argv) == 2:
        company = str(sys.argv[1])

else:
    print("Usage: guvscraper.py 'companyname'")
    sys.exit()

def get_guv(company):
    
    URL = base_url + company + ""
    print(URL)
    
    page = requests.get(URL, headers=user_agent)
    soup = BeautifulSoup(page.content, "html.parser")
    name = soup.find("h2", {"class":"font-resize"} ).get_text()
    boxTableList = soup.findAll('div', attrs={"class" : "box table-quotes"})
    headlineList = soup.findAll('h2', attrs={"class" : "box-headline"})

    #Export to HTML
    #with open(company + ".html", "w", encoding='utf-8') as file:
    #    file.write(str(boxTableList))
    
    print(name+"\n")
    #print(boxTableList)
    
    dflist = pd.read_html(str(boxTableList), decimal=',', thousands='.')

    print("Writing to .xlsx.....")
    write_to_xlsx(dflist, company, headlineList, name)


def write_to_xlsx(dataframelist, company, headlines, name):

    headlinerow = 3
    row = 3
    spacing = 3
    
    #Setup excel file and formatting
    writer = pd.ExcelWriter(company + "_" + str(datetime.date(datetime.now())) + ".xlsx",engine='xlsxwriter', options={'strings_to_numbers': True})   
    workbook=writer.book
    worksheet_name = company + " Bilanzen"
    worksheet=workbook.add_worksheet(worksheet_name)
    writer.sheets[worksheet_name] = worksheet
    bold = workbook.add_format({'bold': True})

    #write company headline
    worksheet.write(1, 0, name, bold)

    for x in range(len(dataframelist)):

        #write GUV data
        dataframelist[x] = dataframelist[x].drop(dataframelist[x].columns[0], 1)
        dataframelist[x].to_excel(writer,sheet_name=worksheet_name,startrow=row , startcol=0, index = False )   
        row += (len(dataframelist[x]) + spacing)

        #write headlines
        worksheet.write(headlinerow, 0, str(headlines[x].get_text()), bold)
        headlinerow += (len(dataframelist[x]) + spacing)

    writer.save()
        

get_guv(company)