import requests
import json
from html.parser import HTMLParser
from bs4 import BeautifulSoup
from html.parser import HTMLParser

url = 'https://www.investing.com/technical/Service/GetSummaryTable'
payload = {
    "tab" : "userQuotes",
    "options[periods][0]": 60,
    "options[periods][1]": 300,
    "options[periods][2]": 900,
    "options[periods][3]": 3600,
    "options[email_hour]": 17,
    "options[timezone_offset]": 19800,
    "options[currencies][]": 166
}
# Adding empty header as parameters are being sent in payload
headers = {
    "Accept":"application/json, text/javascript, */*; q=0.01",
    "Accept-Encoding":"gzip, deflate, br",
    "Accept-Language":"en-IN,en;q=0.9,kn-IN;q=0.8,kn;q=0.7,en-US;q=0.6",
    "Connection":"keep-alive",
    "Content-Length":202,
    "Content-Type":"application/x-www-form-urlencoded",
    "Host":"www.investing.com",
    "Origin":"https://www.investing.com",
    "Referer":"https://www.investing.com/technical/technical-summary",
    "Sec-Fetch-Dest":"empty",
    "Sec-Fetch-Mode":"cors",
    "Sec-Fetch-Site":"same-origin",
    "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36",
    "X-Requested-With":"XMLHttpRequest"
}
print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
r = requests.post(url, data=payload, headers=headers)
# print(r.content)

# print("JUST R> CONTENT ENDS HERRE ..............................")
# print("Print each key-value pair from JSON response")
# for key, value in r.json().items():
#     print(key, ":", value)
# #print(r.content)
# print("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<")

print (r.json()["html"])

tableHead =[]
tableBody =[]

soup = BeautifulSoup(r.json()["html"], "html.parser")
print("asdasdadsad")
for string in soup.thead.stripped_strings:
    tableHead.append(string)

for string in soup.tbody.stripped_strings:
    tableBody.append(string)

print(tableBody[0])
tableHead = tableHead [2:]
tableBody = tableBody [2:]

obj = [{
    tableHead[0]:{
    tableBody[0] : tableBody[1],
    tableBody[5] : tableBody[6],
    tableBody[10]  : tableBody[11]
}},
{  
    tableHead[1]:{
    tableBody[0] : tableBody[2],
    tableBody[5] : tableBody[7],
    tableBody[10]  : tableBody[12]
}},
{  
    tableHead[2]:{
    tableBody[0] : tableBody[3],
    tableBody[5] : tableBody[8],
    tableBody[10]  : tableBody[13]
}},
{  
    tableHead[3]:{
    tableBody[0] : tableBody[4],
    tableBody[5] : tableBody[9],
    tableBody[10]  : tableBody[14]
}}
]

print(obj)
#print(obj[3]['Daily']['Indicators:'])