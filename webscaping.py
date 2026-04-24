from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as Ec
from openpyxl import Workbook    


sheet_location=r'C:/Users/Admins/Downloads/WebScrabed.xlsx'
xpath="//*[contains(text(),'Published')]"
datetext="Published:"

url='https://www.aiimsexams.ac.in/'

wb=Workbook()
ws=wb.active
ws.title='Scrabed Data'

def storedata(**args):
    for result in args:
        ws.append(result)
    wb.save(sheet_location)

driver=webdriver.Chrome()
driver.get(url)
driver.set_page_load_timeout(30)
driver.maximize_window()
WebDriverWait(driver,30).until(Ec.presence_of_element_located((By.XPATH, xpath)))

#page title
print(driver.title)
html=driver.page_source  


#Find all links the website
links=driver.find_elements(By.TAG_NAME, 'a')
print("Total number of links........",len(links))
for lnk in links:
    objlink=storedata()
    objlink(lnk.text, lnk.get_attribute('href'))


#Getting published date
def publisheddate(xpath, datetext):
    pubdates=driver.find_elements(By.XPATH, xpath)
    print(type(pubdates))
    for a in pubdates:
        storedata(a.text.split( )[1], a.text.split(datetext)[-1].strip())
      
publisheddate(xpath, datetext)