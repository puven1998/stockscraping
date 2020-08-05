from selenium import webdriver
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import time
import openpyxl
import pandas as pd

allsectors = {1:'BOND ISLAMIC',2:'CLOSED END FUND',3:'CONSTRUCTION',4:'CONSUMER PRODUCTS & SERVICES',5:'ENERGY',6:'EXCHANGE TRADED FUND-BOND',
              7:'EXCHANGE TRADED FUND-COMMODITY',8:'EXCHANGE TRADED FUND-EQUITY',9:'FINANCIAL SERVICES',10:'HEALTH CARE',
              11:'INDUSTRIAL PRODUCTS & SERVICES',12:'PLANTATION',13:'PROPERTY',14:'REAL ESTATE INVESTMENT TRUSTS',
              15:'SPECIAL PURPOSE ACQUISITION COMPANY',16:'TECHNOLOGY',17:'TELECOMMUNICATIONS & MEDIA',
              18:'TRANSPORTATION & LOGISTICS',19:'UTILITIES'}
dfs = {}
driver = webdriver.Chrome()
driver.get("https://www.malaysiastock.biz/Listed-Companies.aspx?type=S&s1=4")
driver.refresh()
for i in allsectors:
    sectorchosen = int(i)
    sector = Select(driver.find_element_by_id('ddlSecuritySector'))
    sector.select_by_visible_text(allsectors[sectorchosen])
    subsector = Select(driver.find_element_by_id('ddlSecuritySubSector'))
    allsubsectors = subsector.options
    subsectors =[]
    for option in allsubsectors:
        subsectors.append(option.text)
    subsectors.remove('--- Filter By Sub Sector ---')
    subsectors.remove('All Sub Sectors')
    stockquote = ['Stockquote']
    market = ['Market']
    compname = ['Name']
    marketcap = ['Market Cap']
    shareprice = ['Latest Share Price(5mins)']
    PE = ['P/E']
    DY = ['D/Y']
    ROE = ['ROE']
    ssubsector = ['Subsector']
    for chosensubsec in subsectors:
        sector = Select(driver.find_element_by_id('ddlSecuritySector'))
        sector.select_by_visible_text(allsectors[sectorchosen])
        subsector = Select(driver.find_element_by_id('ddlSecuritySubSector'))
        subsector.select_by_visible_text(chosensubsec)
        print('You have chosen',chosensubsec,'in',allsectors[sectorchosen],'sector')
        time.sleep(1)
        source = driver.page_source
        soup = BeautifulSoup(source,'lxml')
        table = soup.find('table',class_='marketWatch')
        companies = table.find_all("tr")[1:]
        for comp in companies:
            try:
                stockquote.append(comp.find('td').find('h3').find('a').text)
            except:
                stockquote.append(None)
            try:
                market.append(comp.find('td').find('span').text)
            except:
                market.append(None)
            try:
                compname.append(comp.find('td').findAll('h3')[1].text)
            except:
                compname.append(None)
            try:
                marketcap.append(comp.findAll('td')[3].text)
            except:
                marketcap.append(None)
            try:
                shareprice.append(float(comp.findAll('td')[4].text))
            except:
                shareprice.append(None)
            try:
                PE.append(comp.findAll('td')[5].text)
            except:
                PE.append(None)
            try:
                DY.append(float(comp.findAll('td')[6].text))
            except:
                DY.append(None)
            try:
                ROE.append(comp.findAll('td')[7].text)
            except:
                ROE.append(None)
            ssubsector.append(chosensubsec)
    dataframe = [compname,ssubsector, stockquote, market, shareprice, marketcap, PE, DY, ROE]
    df = pd.DataFrame(dataframe)
    df1 = df.transpose()
    df.rename(columns = df.iloc[0],inplace=True)
    df.drop([0],inplace=True)
    dfs[allsectors[sectorchosen]] = df1

writer = pd.ExcelWriter('Valueinvesting1.xlsx')
for sheetname in dfs.keys():
    if sheetname == 'SPECIAL PURPOSE ACQUISITION COMPANY':
        dfs[sheetname].to_excel(writer, sheet_name=sheetname[0:27], index=False)
    else:
        dfs[sheetname].to_excel(writer,sheet_name=sheetname,index=False)
writer.save()
writer.close()
driver.quit()
