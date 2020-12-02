#importing fpdf module
from fpdf import FPDF

# save FPDF() class into  a variable pdf
pdf = FPDF()

# Add a page into the pdf created
pdf.add_page()
   
# set style and size of font that you want in the pdf 
pdf.set_font("Arial", size = 5.3)

# open the json file in read mode 
f = open("file.json", "r")

# insert the texts from json file  in pdf
for x in f:
    pdf.cell(100, 6, txt = x, ln = 1, align = 'A')
    
# saves the pdf with name .pdf 
pdf.output("MyProject.pdf")
