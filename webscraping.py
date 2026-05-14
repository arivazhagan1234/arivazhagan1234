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
allproductsxpath="//*[contains(@class,'a-section a-spacing-base desktop-grid-content-view')]"
productlabelxpath=".//h2"
ratingxpath=".//*[contains(@class,'a-row a-size-small')]"
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

    newrow=df([row], columns=['Page Title', 'Product Label', 'PRODUCT RATING','LinkText', 'URL'])
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
pageurl=[]
productdata=[]

#Find all links the website
links=driver.find_elements(By.TAG_NAME, 'a')
print("Total number of links........",len(links))
def pagelinks(links):
    for lnk in links:    
        linktext=lnk.text
        linkurl=lnk.get_attribute('href')
        if linkurl and (linkurl.startswith('http://') or linkurl.startswith('https://')):
            print("inside of if Treue condition........", linktext, linkurl)
            pageurl.append((linkurl, linktext))        
pagelinks(links)    

#to get product label and discription
def productsdescription(allproductsxpath,productlabelxpath,ratingxpath):

    allproducts=wait.until(Ec.presence_of_all_elements_located((By.XPATH, allproductsxpath)))
    for product in allproducts:
        try:
            label=product.find_element(By.XPATH,productlabelxpath).text.strip()   
        except TimeoutException:
            label="No label found"
           
        try:
            rating=product.find_element(By.XPATH,ratingxpath).text.strip()     
        except TimeoutException:
            rating="No rating"

        productdata.append((label, rating))
        print("Label:", label, "| Rating:", rating)
    return productdata
productsdescription(allproductsxpath,productlabelxpath,ratingxpath)

'''  
def productlabel(productlabelxpath):
    productlabels=wait.until(Ec.presence_of_all_elements_located((By.XPATH, productlabelxpath)))
    label=[l.text.strip() for l in productlabels if l.text.strip() != '']
    return label

#to get product rating
def productrating(ratingxpath):
    rating=wait.until(Ec.presence_of_all_elements_located((By.XPATH, ratingxpath)))
    productratings=[r.text for r in rating if r.text != '']
    return productratings

'''  
print(pageurl)
#store page title,url,urltext
def allpageurl(pageurl,allproductsxpath,productlabelxpath,ratingxpath):
    for linkurl, linktext in pageurl:
        try:
            driver.get(linkurl)
            pagetitle=driver.title
            productsdes=productsdescription(allproductsxpath,productlabelxpath,ratingxpath)
            for label,rating in productsdes:  
                storedata(pagetitle, label, rating, linktext, linkurl)
        except TimeoutException:
            print("Time out url is ......",linkurl)
            driver.execute_script("window.stop();")
allpageurl(pageurl,allproductsxpath,productlabelxpath,ratingxpath)
                                                                                   

