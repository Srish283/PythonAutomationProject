import requests
from bs4 import BeautifulSoup
import csv

URL = "http://www.values.com/inspirational-quotes"
r = requests.get(URL)
try:
 r.raise_for_status()
except Exception as exc:
 print('There was a problem: %s' % (exc))
soup = BeautifulSoup(r.content, 'html5lib')

quotes=[]  # a list to store quotes

table = soup.find('div', attrs = {'id':'all_quotes'})

for row in table.findAll('div',
                         attrs = {'class':'col-6 col-lg-3 text-center margin-30px-bottom sm-margin-30px-top'}):
    quote = {}
    quote['theme'] = row.h5.text
    quote['url'] = row.a['href']
    quote['img'] = row.img['src']
    quotes.append(quote)

filename = 'inspirational_quote.csv'
with open(filename, 'w', newline='') as f:
    w = csv.DictWriter(f,['theme','url','img'])
    w.writeheader()
    for quote in quotes:
        w.writerow(quote)
