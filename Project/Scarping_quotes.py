#-----Webscraping  and storing the information in the .csv file--------


import requests
from bs4 import BeautifulSoup
import csv
import json
from fpdf import FPDF
print('Mini Project to Demonstrate Working of following Concepts:  ')
print("---------------------------------------------------------------------")

print('1.Webscraping the data')
print('2.Storing the scraped data in the .csv file')
print('3.Converting the .csv file to .json file')
print('4.Converting the .json file to pdf')

print("---------------------------------------------------------------------")
def open_url():

    #Enter the url from where u want to search the quotes
    #URL = "http://www.values.com/inspirational-quotes"
    print("Lets webscrap the data from the URL and store it in .csv file")
    URL = "http://www.values.com/inspirational-quotes"
    r = requests.get(URL)
    print('--- URL is entered by the user ---')

    try:
        r.raise_for_status() #checks for error
    except Exception as exc:
        print('There was a problem: %s' % (exc))
    soup = BeautifulSoup(r.content, 'html5lib')

    quotes=[]  # a list to store quotes

    table = soup.find('div', attrs = {'id':'all_quotes'}) # finds attributes where id is all_quotes

    for row in table.findAll('div',
                            attrs = {'class':'col-6 col-lg-3 text-center margin-30px-bottom sm-margin-30px-top'}):
        quote = {}
        quote['theme'] = row.h5.text
        quote['url'] = row.a['href']
        quote['img'] = row.img['src']
        quotes.append(quote)

    filename = 'inspirational_quote.csv' # filename is inspirational_quote.csv
    with open(filename, 'w', newline='') as f:  # opens the file in write mode
        w = csv.DictWriter(f,['theme','url','img'])
        w.writeheader() # writes the header (theme,url,img)

        for quote in quotes:
            w.writerow(quote) # writes the quotes
    print('Operation sucessfull')

    print('---Quotes successfully stored in csv file---')


#Converting .csv file  to .json file

def csv_to_json():
    print("---------------------------------------------------------------------")
    print("Lets convert the csv file to json file")
    csvfile = open('inspirational_quote.csv', 'r') # opens the file in read mode
    print('-----csv file is opened!-----')

    jsonfile = open('file.json', 'w') # opens the file in write mode
    print('-----New json file is created-----')

    fieldnames = ('theme','url','img')

    reader = csv.DictReader( csvfile, fieldnames) # read file using DictReader
    print('-----All the fields in the csv file are stored in json Successfully-----')

    for row in reader:
        json.dump(row, jsonfile) #converts the Python objects into appropriate json objects
        jsonfile.write('\n')    #writes

    print('-----Successfully converted csv file to json file-----')


#Converting .json file to .pdf file


def json_to_pdf():
    print("---------------------------------------------------------------------")
    print("Lets convert the json file to pdf file")

    pdf = FPDF() # save FPDF() class into a variable pdf
    print('-----pdf is being created!-----')
    pdf.add_page() # Add a page

    print('-----A Page is created in the pdf!-----')

    pdf.set_font("Arial", size = 5.3) # set style and size of font that you want in the pdf

    f = open("file.json", "r") # open the text file in read mode

    for x in f:
        pdf.cell(0, 10, txt= x, ln = 2, align='A') # insert the texts in pdf
        #print(x)

    pdf.output("MyProject.pdf") # save the pdf with name .pdf

    print('-----Converted a json file to pdf successfully-----')

    print("---------------------------------------------------------------------")

if __name__=="__main__": #Function Calls
    open_url()

    csv_to_json()

    json_to_pdf()


    

    

