from fpdf import FPDF
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size = 5.3)
f = open("file.json", "r")
for x in f:
    pdf.cell(100, 6, txt = x, ln = 1, align = 'A')
pdf.output("mygfg.pdf")
