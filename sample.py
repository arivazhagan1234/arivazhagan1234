
"""
def add(a,b):
    return a+b

print("The addition",add(2,4)
"""
from wsgiref import headers

import requests
import certifi  
import lxml
from bs4 import BeautifulSoup
from urllib.parse import urlparse, urljoin, urlunparse

"""

from urllib.parse  import urlparse, urljoin, urlunparse
base="https://www.w3schools.com/"

add={'name':{'name':'Ariva','url': '//example//eneva//?id=23#hareff','gender':'male'}}
add1={'Name':'Ariva','age':23,'gender':'male'}
add2={'Studentname':'Ariva','age':23,'gender':'male'}


def urlremax(url, base):
    if not url:
        print("The url is empty")
        return ""
    if isinstance(url, str):
        print("The url is string",url)
        url = url.strip()
        return urljoin(base, url)
    if isinstance(url, dict):
        url =url.get('url', '')
        if url.startswith('http'):
            url = urljoin(base, url.replace('http', 'https'))
            return url
        if url.startswith('/') or url.startswith('//'):
            url = urlparse(url.lstrip('/'))
            url1 = urlunparse((url.scheme, url.netloc, url.path, "", "", ""))
            url2 = urljoin(base, url1)    
           
            return url2

def content(add1, *keys, default=None):
    print(type(add1))
    for key in keys:
        
        if key in add1: 
            data = add1.get(key)
            return data
    low={str(k).lower(): k for k in add1.keys() }
    for key in keys:
        if key.lower() in low:
            data = add1.get(low[key.lower()]); url=urlremax(data, base)
            return url

    return default
print(content(add, 'Name','Custname','stuadentname'))




payload={'search': "angola luanda", 'top': '10', 'skip': '0', 'count': True, 'queryType': 'simple', 'filter': 'content/TransactionTypeUID eq 260'}

url='https://www.remax-multitrust.co.ao/search/listing-search/docs/search'
ses=requests.Session()
ses.headers.update({'User-Agent': 'Mozilla/5.0 (windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'})
r=ses.post(url, json=payload, headers={'Content-Type': 'application/json'})
print(r.status_code)
print(r.headers.get('Content-Type'))
print(r.url)
data=r.json()
#print(data.get('value'))
print( "the length of data.get('value') is:", len(data.get('value')))
#print(data['value'][0])
for item in data['value']:
    print("the item type is:", type(item))
    content=item['content']
    print(content.get('TitleAddress'))
    print(content.get('City'))


url='https://angoimoveis.com/category/venda'
ses=requests.Session()
ses.headers.update({'User-Agent': 'Mozilla/5.0 (windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'})
response=ses.get(url)
print(response.status_code)
soup=BeautifulSoup(response.text, "lxml")
print(soup.prettify())
#print(soup.select_one('City'))
"""
payload={'search': "angola luanda", 'top': '10', 'skip': '0', 'count': True, 'queryType': 'simple', 'filter': 'content/TransactionTypeUID eq 260'}

url='https://www.remax-multitrust.co.ao/search/listing-search/docs/search'

header={'Authorization' : 'Bearer 2c7e3f0b-1d4e-4a5f-8c9b-6e5f3d2e1a4b'}
response = requests.get(url, payload=payload, headers=header)
print(response.status_code)