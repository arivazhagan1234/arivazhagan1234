from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as Ec
from selenium.common.exceptions import TimeoutException
from openpyxl import Workbook
from openpyxl import load_workbook
from pandas import DataFrame as df
import os
import time

sheet_location=r'C:/Users/Admins/Downloads/WebScrabed.xlsx'
csv_location=r'C:/Users/Admins/Downloads/WebScrabed.csv'
css01='h2'
css02='h1'
xpath1="//*[contains(text(),'Published')]"
xpath2="Published:"
productlabelxpath="//h2"
url='https://www.amazon.in/'

#creat a new excel or reuse existing exel sheet
def storedata(*args):
    row=list(args)
    print("type of result....", row)  
    if not os.path.exists(sheet_location):
        wb=Workbook()
        ws=wb.active
        ws.title='Web scraped'
        ws.append(row)
    else:
        wb=load_workbook(sheet_location)
        ws=wb.active
    ws.append(row)
    wb.save(sheet_location)

    newrow=df([row], columns=['Page Title', 'LinkText', 'URL'])
    #create new csv file if not have existing file or use existing file  
    if os.path.exists(csv_location):
        newrow.to_csv(csv_location, mode='a', header=False, index=False)
    else:
        newrow.to_csv(csv_location, mode='w',header=True, index=False)
        


#create new browser instance
driver=webdriver.Chrome()
driver.get(url)
driver.set_page_load_timeout(30)
driver.maximize_window()
driver.implicitly_wait(30)
driver.set_script_timeout(30)
wait=WebDriverWait(driver,30)
#wait.until(Ec.presence_of_element_located((By.XPATH, xpath1)))

#page title
print(driver.title) 

#Store the all urls
pagedata=[]
productlabels=[]

#Find all links the website
links=driver.find_elements(By.TAG_NAME, 'a')
print("Total number of links........",len(links))
def pagelinks(links):
    for lnk in links:    
        linktext=lnk.text
        linkurl=lnk.get_attribute('href')
        if linkurl and (linkurl.startswith('http://') or linkurl.startswith('https://')):
            print("inside of if Treue condition........", linktext, linkurl)
            pagedata.append((linkurl, linktext))        
pagelinks(links)    

def productlabel(productlabelxpath):
    productlabels=wait.until(Ec.presence_of_all_elements_located((By.XPATH, productlabelxpath)))
    for a in productlabels:
        label=a.text.strip()
        productlabels.append(label)

print(pagedata)
#store page title,url,urltext
def allpagedata(pagedata, productlabelxpath):
    for linkurl, linktext in pagedata:
        try:
            driver.get(linkurl)
            pagetitle=driver.title
            productlabel(productlabelxpath)
            storedata(pagetitle,linkurl,linktext)
        except TimeoutException:
            print("Time out url is ......",linkurl)
            driver.execute_script("window.stop();")
allpagedata(pagedata, productlabelxpath)



print("productlabels.......................", productlabels)


"""
#Getting published date
def publisheddate(xpath1, xpath2):
    pubdates=driver.find_elements(By.XPATH, xpath1)
    print(type(pubdates))
    for a in pubdates:
        textword=a.t   ext.split( )[1]
        dates=a.text.split(xpath2)[-1].strip()
        storedata(textword, dates)
      
publisheddate(xpath1, xpath2)
"""