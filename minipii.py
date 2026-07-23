from urllib.parse import urlparse, urlunparse, urljoin
import requests
import time
from pathlib import Path
import re
from bs4 import BeautifulSoup
import logging


PROJECTROOT=Path.cwd()
logging.basicConfig(level=logging.INFO, format="%(asctime)s-%(levelname)s-%(filename)s-%(message)s")

base="https://angoimoveis.com/"


urls=["//fonts.gstatic.com","/cdn-cgi/l/email-protection#452c2b232a05242b222a2c282a33202c366b262a28",'https://angoimoveis.com/search','https://angoimoveis.com/search?l=71&amp;c=16','https://www.Angoimoveis.com/category/venda?page=2', 'https://angoimoveis.com/vivenda-v6-anexo-vila-alice-luanda-25197', 'https://angoimoveis.com/terreno-60-30-zango-8000-25199']

def urlreform(url, base):
    ex=('angoimoveis.com', 'venda')
    if not url or not isinstance(url, str) or not any(word in url for word in ex):
        return ""
    url=str(url).strip()

    logging.warning("Started abslute url")
    if url.startswith('/'):
        url=urljoin(base, url.lstrip('/'))
    if url.startswith('http://'): 
        url= url.replace('http://', 'https://', 1)
    parse=urlparse(url.strip())
    absurls=urlunparse((parse.scheme, parse.netloc, parse.path, "", parse.query,""))
    logging.warning("created abslute url successfully")
    return absurls
            
def baseurl(url):
    try:
        return urlparse(url).netloc.lower().replace('www.', '')
    except Exception:
        print("The exception", url )
        return ""

CNFG: dict= {"SAVE_HTML": True, "TIME_OUT":20, "MAX_RETRY":4, "DEBUG": PROJECTROOT / "debug_html"}

def ses():
    s=requests.Session()
    s.headers.update({"User-Agent":(
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " "AppleWebKit/537.36 (KHTML, like Gecko)"  "Chrome/150.0.0.0 Safari/537.36"),
        "Acceptence": "application/json text/plain */*",
        "Content_Type": "application/json",
        "Accept-Language": "pt-PT,pt;q=0.9,en;q=0.8",
        "Origin": base})
    return s


def _get(s,url):
    for t in range(1,CNFG["MAX_RETRY"]+1):
              #print("in range Url r ,,,,,,,,,", r)
       try:
          r=s.get(url, timeout=CNFG["TIME_OUT"])
          #raise requests.exceptions.RequestException ("this is test exception")
          r.raise_for_status()
          if CNFG["SAVE_HTML"]== True:
             d=CNFG["DEBUG"]; d.mkdir(parents=True, exist_ok=True)
             fpt=re.sub(r"[^\w]", ' ', url)
             (d/f'{fpt}.html').write_text(r.text, 'utf-8')    
          return BeautifulSoup(r.text, 'lxml')
       except requests.exceptions.HTTPError as e:
            print("inside the https error")
            if e.response is not None and e.response.status_code==404:
               logging.warning(f'(Angoimoies) HTTP {e} | tentative: {t}')
               return None   
       except requests.exceptions.RequestException as e:
          logging.warning(f'Angomoives {e} | tentative: {t}')
       logging.warning(f"the tentetive is ..........................{t}")
       if t <= CNFG["MAX_RETRY"]: time.sleep(5*t) 
    return None        
          
for url in urls:
    #url='https://angoimoveis.com/search?l=71&amp;c=1686'
    reformedurl=[]
    absurls=urlreform(url, base)
    if absurls not in ("", None): reformedurl.append(absurls)
    reformedurl=list(dict.fromkeys(reformedurl))
    print("reforedurl,,,,,,,,,,,,,,", reformedurl)
    for url in reformedurl:
        s=ses()
        soup=_get(s,absurls)
        print("the value url............",absurls)
        print(soup.select_one('h1'))

 