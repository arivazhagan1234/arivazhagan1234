from urllib.parse import urlparse, urlunparse, urljoin
import requests
import time, random
from pathlib import Path
import re
from bs4 import BeautifulSoup
import logging


PROJECTROOT=Path.cwd()
logging.basicConfig(level=logging.INFO, format="%(asctime)s-%(levelname)s-%(filename)s-%(message)s")


base="https://angoimoveis.com/"
CNFG: dict= {"SAVE_HTML": True, "TIME_OUT":20, "MAX_RETRY":4, "DEBUG": PROJECTROOT / "debug_html", "DELAY_MIN":3, "DELAY_MAX":5, "TEXTLIMIT":1000}


urls=['https://angoimoveis.com/search','https://angoimoveis.com/search?l=71&amp;c=16','https://www.Angoimoveis.com/category/venda?page=2', 'https://angoimoveis.com/vivenda-v6-anexo-vila-alice-luanda-25197', 'https://angoimoveis.com/terreno-60-30-zango-8000-25199']

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
              
       try:
          r=s.get(url, timeout=CNFG["TIME_OUT"])
          #raise requests.exceptions.RequestException ("this is test exception")
          logging.info(f'the status code is {r.status_code}')
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
          
def delay():
    time.sleep(random.uniform(CNFG['DELAY_MIN', CNFG['DELAY_MAX']]))


inputstr="Imóveis a  venda e imóveis para alugar. Na                     Angoimoveis você encontra anúncios               classificados de imóveis para compra, venda ou aluguel em Angola. óveis a venda e imóveis para alugar. Na Angoimoveis você encontra anúncios classificados de imóveis para compra, venda ou aluguel em Angola."

def reformtext(text, Textlimit=None):
   
    if not isinstance(text, str) or text in ("", None):
        return
    s=re.sub(r'\s+', ' ', text)
    return s[:Textlimit]  if s not in ("", None) else s


def contentreform(cont):
    textcont=[]
    
    def work(cont):
        if len(' '.join(textcont)) > CNFG["TEXTLIMIT"]:
                return
        if isinstance(cont, dict):
            for k, v in cont.itens():
                lk=k.lower()
                if lk in ("url", "logo", "sameAs", "shortLinks", "ShortLinks"): continue
                print("content is .............", k)
                work(v)
        elif isinstance(cont, list):

            for val in cont:
                if val in ("url", "logo", "sameAs", "shortLinks", "ShortLinks"): continue
                print("content is .............", val)
                work(val)
        elif isinstance(cont, (str, float, "")):
            newfornmtext = reformtext(cont, 200)
            textcont.append(newfornmtext)
    work(cont)
        
    return " ".join(textcont) or " "

reformedurl=[]
for url in urls:
    absurls=urlreform(url, base)
    if absurls not in ("", None): reformedurl.append(absurls)
    reformedurl=list(dict.fromkeys(reformedurl))

print( " reformurl", reformedurl)
for url in reformedurl:
    print("Theemabsbsnbas", url)
    s=ses()
    soup=_get(s, url)
    if soup in ("", None):
        continue
    
    h1=soup.select_one('h1') or soup.select_one('h2')

    print("the hi tags....", h1)
    resultext=contentreform(h1.get_text(" ", strip=True)) if h1 else urlparse.path.strip('/').splite('/').replace("-", " ").title()
    print(" the resulttext is", resultext)
    ("resultext............",h1.get_text("", strip = True))
     soup
