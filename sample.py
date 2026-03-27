
"""
def add(a,b):
    return a+b

print("The addition",add(2,4)
"""
import requests
import certifi  

urls=['https://www.google.com','https://www.example.com/', 'https://www.jimdo.com/examples/']

for url in urls:
    response=requests.get(url, verify=False)
    if response.status_code==200:
        print(url, response.status_code)
    else:
        print(url, "Website is Down")



