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
price_xpath=".//span[@class='a-price']/span[@class='a-offscreen']"
mrp_xpath=".//span[contains(@class,'a-text-price')]/span[@class='a-offscreen']"
discount_xpath=".//span[contains(text(),'off')]"

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

    newrow=df([row], columns=['Page Title', 'Product Label', 'PRODUCT RATING','PRICE','MRP','DISCOUNT', 'URL'])
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
scraping=["https://www.amazon.in/computers-and-accessories/b/?ie=UTF8&node=976392031&ref_=nav_cs_pc","https://www.amazon.in/s/ref=mega_elec_s23_2_2_1_1?rh=i%3Acomputers%2Cn%3A1375458031&ie=UTF8&bbn=976392031","https://www.amazon.in/gift-card-store/b/?ie=UTF8&node=3704982031&ref_=nav_cs_gc"]

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
def productsdescription(allproductsxpath,productlabelxpath,ratingxpath,price_xpath, mrp_xpath, discount_xpath):

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

        try:
            price = product.find_element(By.XPATH, price_xpath).text
        except:
            price = "No price found"
        
        try:
            mrp = product.find_element(By.XPATH, mrp_xpath).text
        except: 
            mrp = "No MRP found"

        try:
            discount = product.find_element(By.XPATH, discount_xpath).text
        except:
            discount = "No discount found"

        productdata.append((label, rating, price, mrp, discount))
        print("Label:", label, "| Rating:", rating, "| Price:", price, "| MRP:", mrp, "| Discount:", discount)
    return productdata

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
def allpageurl(scraping,allproductsxpath,productlabelxpath,ratingxpath,price_xpath, mrp_xpath, discount_xpath):

    for url in scraping:
        try:
            driver.get(url)
            time.sleep(4)
            pagetitle=driver.title
            print("page url is..........................", url)
            productsdes=productsdescription(allproductsxpath,productlabelxpath,ratingxpath,price_xpath, mrp_xpath, discount_xpath)
            for label,rating,price,mrp,discount in productsdes:  
                storedata(pagetitle, label, rating, price,mrp,discount, url)
        except TimeoutException:
            print("Time out url is ......",url)
            driver.execute_script("window.stop();")
allpageurl(scraping,allproductsxpath,productlabelxpath,ratingxpath,price_xpath, mrp_xpath, discount_xpath)
                                                                                   

