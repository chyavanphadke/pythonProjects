import datetime
from urllib.request import Request, urlopen
from bs4 import BeautifulSoup as soup
url = 'https://www.investing.com/indices/us-spx-500-futures-technical'
req = Request(url , headers={'User-Agent': 'Mozilla/5.0'})
import winsound

freq = 2500
dur = 400

neutralCount = 0

def getData():
    global title
    global now
    webpage = urlopen(req).read()
    page_soup = soup(webpage, "html.parser")
    title = page_soup.find_all("div", {'class':'newTechStudiesRight instrumentTechTab'})[0].find('span').text
    now = datetime.datetime.now()

while (True):
    getData()
    if (title == 'Neutral'):
        print(now.strftime("%H:%M:%S"), " Wait, its NEUTRAL")
        winsound.Beep(freq,dur)
        neutralCount = neutralCount + 1
    else:
        print(now.strftime("%H:%M:%S"), title, " Sold ")
        neutralCount = 0

    while (neutralCount > 20):
        getData()
        if(title != 'Neutral'):
            print("Stable and good to go with ", title, " at time = ", now )
            dur = 800
            winsound.Beep(freq,dur)
            neutralCount = 0
        if(title == 'Neutral'):
            print("Wait for decision")