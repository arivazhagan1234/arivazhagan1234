from os import name
import os
import openpyxl

import requests
from bs4 import BeautifulSoup
import time
from selenium import webdriver   
import sqlite3
url = "https://www.remax-multitrust.co.ao/"
"""  
def get_remax_listings(url):
    driver=webdriver.Chrome()
    driver.get(url)
    driver.maximize_window()
    time.sleep(7)
    html=driver.page_source
    soup = BeautifulSoup(html, "html.parser")
   
    h1_tags = soup.find_all("h1")

    print("The h1 tags are:", h1_tags)
    print("Number of h1 tags:", len(h1_tags))

    titles = []

    for tag in h1_tags:
        titles.append(tag.get_text(strip=True))

    return titles
]-.?: d.ēA.]/

titles = get_remax_listings(url)

con=sqlite3.connect("mybio.db")
con.row_factory= sqlite3.Row

curser=con.cursor()

def get_table_val(qinfo=None):
    q="select * from bio where gender='male'"; p=[]
    if qinfo:
        print("Query Info:",qinfo)
        for name in ('name','age','telno'):
            if qinfo.get(name):q+=f" AND {name}=?"; p.append(qinfo[name])
        
    print("Query:",q)
    print("Params:",p)

    curser.execute(q,p)
    return [dict(r) for r in curser.fetchall()]
result=get_table_val()
con.close()
print(result)
"""
con=sqlite3.connect("mybio.db")
con.row_factory=sqlite3.Row
cursor=con.cursor()
def get_table_val(qinfo=None):
    q="select * from bio;"
    c=cursor.execute(q)
    return [dict(r) for r in c.fetchall()]
wb=openpyxl.Workbook()
root=os.path.join(os.getcwd(), "mybio.xlsx")
wb.create_sheet(title="bio",index=0)
ws=wb.active

columns=["name","age","gender","cell_num",'Address','Email'] 
ws.append(columns)
val=get_table_val()
for v in val:
    if type(v) is dict:
     ws.append([v.get(c) if v.get(c) else "" for c in columns] )

wb.save(root)
con.close()










