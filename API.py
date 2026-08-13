
import requests, json
import argparse 
import re


parser = argparse.ArgumentParser(description = "Addiction API Client")
parser.add_argument("--val1", type = str, required = True, help = "First number of addition operation")
parser.add_argument("--val2", type = int , required = True, help = "Second number of addition operation")
parser.add_argument("--val3", type = str , required = True, help = "Third number of addition operation")
args = parser.parse_args()

urls= ["http://127.0.0.1:5000/addition","http://127.0.0.1:5000/multiplication"]
url = "http://127.0.0.1:5000/addcontact"

payload = {"val1" : args.val1, "val2" : args.val2, "val3": args.val3}
headers = {"Content-Type" : "application/json"}
"""
def fun(url):
    try:
        response = requests.post(url, json = payload, headers = headers)
        if response.status_code == 200: 
            print("Addition Result:", response.json())
    except requests.exceptions.RequestException as e:
        print("Error occurred:", e)

def urlreform(url):
    
    #if not url or not url.startswith("http://") or not url.startswith("https://"):
     #   return "" 

    for txt in ("addition","multiplication"):
        txt = txt.strip()
        match = re.search(txt, url)
        if match:
            print(" the txt", txt)
            return url, txt
        
    return url

for url in urls:
    opurl, name = urlreform(url)
    print("thop ,name is " , opurl, name)
    
    if opurl in ("None", " "):
        continue

    if name == "addition":
        result = fun(url)
    elif name == "multiplication":
        result = fun(url) 
    else:
        exit()

"""
def addcontact():
    print("the payload",payload)
    response = requests.post(url, json = payload , headers = headers)
    data = response.json()
    print(data["contact"])
    print(type(data))
    assert response.status_code == 200
    

    #headers = {"authorization": f"Bearer {token}", "Content-Type" : "application/json"}

    #return headers
addcontact()