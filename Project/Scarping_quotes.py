#-----Webscraping  and storing the information in the .csv file--------

import requests
from bs4 import BeautifulSoup
import csv
import json
from fpdf import FPDF 


def open_url():

    #Enter the url from where u want to search the quotes
    #URL = "http://www.values.com/inspirational-quotes" <----- Use this url

    URL = input('Enter a url: ')
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


#Converting .csv file  to .json file

def csv_to_json():
    csvfile = open('inspirational_quote.csv', 'r')

    jsonfile = open('file.json', 'w')

    fieldnames = ('theme','url','img')

    reader = csv.DictReader( csvfile, fieldnames)
    for row in reader:
        json.dump(row, jsonfile)
        jsonfile.write('\n')


#Converting .json file to .pdf file


def json_to_pdf():

    # save FPDF() class into  a variable pdf 
    pdf = FPDF()    
    
    # Add a page into the pdf created
    pdf.add_page() 
    
    # set style and size of font that you want in the pdf 
    pdf.set_font("Arial", size = 5.3) 
    
    # open the .json file in read mode 
    f = open("file.json", "r") 
    
    # insert the texts in pdf created
    for x in f: 
        pdf.cell(0, 10, txt= x, ln = 2, align='A') 
        
        
    # save the pdf with name .pdf 
    pdf.output("MyProject.pdf")   

#Function Calls   
if __name__=="__main__":
    #search for the quotes and store in .csv file
    open_url()

    #Converts inspirational_quote.csv file to file.json file
    csv_to_json()

    #Converts file.json to MyProject.pdf
    json_to_pdf()

    

    

