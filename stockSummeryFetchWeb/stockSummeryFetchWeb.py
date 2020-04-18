import datetime
from urllib.request import Request, urlopen
from bs4 import BeautifulSoup as soup
url = 'https://www.investing.com/indices/us-spx-500-futures-technical'
req = Request(url , headers={'User-Agent': 'Mozilla/5.0'})


while (True):
    
    webpage = urlopen(req).read()
    page_soup = soup(webpage, "html.parser")
    title = page_soup.find_all("div", {'class':'newTechStudiesRight instrumentTechTab'})[0].find('span').text

    now = datetime.datetime.now()
    if (title == 'Neutral'):
        print(now.strftime("%H:%M:%S"), " Wait, its NEUTRAL")
    else:
        print(now.strftime("%H:%M:%S"), " DO this right now :", title)