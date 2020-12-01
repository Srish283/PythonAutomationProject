import csv
import json

csvfile = open('inspirational_quote.csv', 'r')
jsonfile = open('file.json', 'w')

fieldnames = ('theme','url','img')
reader = csv.DictReader( csvfile, fieldnames)
for row in reader:
    json.dump(row, jsonfile)
    jsonfile.write('\n')
