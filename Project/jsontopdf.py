#import fpdf module
from fpdf import FPDF

# save FPDF() class into  a variable pdf
pdf = FPDF()

# Add a page 
pdf.add_page()
   
# set style and size of font that you want in the pdf 
pdf.set_font("Arial", size = 5.3)

# open the text file in read mode 
f = open("file.json", "r")

# insert the texts in pdf
for x in f:
    pdf.cell(100, 6, txt = x, ln = 1, align = 'A')
    
# save the pdf with name .pdf 
pdf.output("MyProject.pdf")
