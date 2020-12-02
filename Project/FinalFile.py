import openpyxl as op
import sqlite3
import PyPDF2
import time
import docx
from openpyxl.styles import Font, Fill
from selenium import webdriver
from docx2pdf import convert


def open_workbook():
    print('---------------------------------------------------------------------------------------------------------------------------------')
    xlsxFile=input('Enter an Excel workbook Name: ')

    try:
        print('OPENING WORKBOOK Employee_Data......')
        wb = op.load_workbook(xlsxFile)  # Open workbook Employee Data i.e Employee_Data.xlsx
        try:
            sheetname=input('Enter Worksheet Name:  ')
            sheet = wb.get_sheet_by_name(sheetname)  # Gets Sheet by sheetname i.e Employee data
            sheet.title='emp_data'                 # Sets title to emp_data
            mysheet=wb.active
            print('Active sheet {}'.format(mysheet))
            print('There are {} rows and {} columns in Employee Data'.format(sheet.max_row,sheet.max_column))   # Highest row and column of the sheet
            sheet['F1']='Job Hours' # Column heading changes from Job_Time to Job Hours
            return sheet,wb
        except KeyError:						# Catches Exception when sheet doesn't exist 
            print('  Worksheet does not exist.')
            exit()

    except Exception as e:
        print("File not found. Check the filename")			# Catches Exception when file doesn't exist and displays error message
        print(e)
        exit()



def create_table():
    print('---------------------------------------------------------------------------------------------------------------------------------')
    print(' Create Table in Database ')
    conn = sqlite3.connect(ConStr)						#Connect to Connection String
    cr=conn.cursor()

    cr.execute('''CREATE TABLE Employee
                 (Emp_id INTEGER PRIMARY KEY, Emp_name TEXT, DOB TEXT,
         Job_Cat INTEGER, Salary INTEGER, Job_Hour TEXT,Project INTEGER,Prev_Exp INTEGER,Gender text)''') # Create table Employee

    print('Table is Created')
    conn.commit()								#Commit Changes
    conn.close()								#Close Connection 


def insert_into_table(sheet):
    print('---------------------------------------------------------------------------------------------------------------------------------')
    print(' Perform Insertion by inserting all Data\'s of Excel to Database ')
    conn = sqlite3.connect(ConStr)						#Connect to Connection String
    cr=conn.cursor()
    for i in range(2, sheet.max_row):         # Iterate for maximum rows
        for j in range(1,sheet.max_column+1): # Iterate for maximum columns
            Emp_id,Emp_name,DOB,Job_Cat,Salary=sheet.cell(row=i,column=1).value,sheet.cell(row=i,column=2).value,sheet.cell(row=i,column=3).value,sheet.cell(row=i,column=4).value,sheet.cell(row=i,column=5).value
            Job_Time,Project,Prev_Exp,Gender=sheet.cell(row=i,column=6).value,sheet.cell(row=i,column=7).value,sheet.cell(row=i,column=8).value,sheet.cell(row=i,column=9).value
            cr.execute('''INSERT OR IGNORE INTO Employee
             (Emp_id,Emp_name,DOB,Job_Cat,Salary,Job_Hour,Project,Prev_Exp,Gender)
            VALUES ( ?, ?, ?, ?, ?, ? ,?,?,?)''',
            ( Emp_id,Emp_name,DOB,Job_Cat,int(Salary),Job_Time,Project,Prev_Exp,Gender)) # Inserts all Excel data into Database ExcelData
    print('Data Inserted into table successfully')
    conn.commit()									#Commit Changes
    conn.close()									#Close Connection 

def delete_rows():
    print('---------------------------------------------------------------------------------------------------------------------------------')
    print(' Perform Deletion on Table by deleting records with Emp id greater then 480')
    conn = sqlite3.connect(ConStr)
    cr=conn.cursor()
    cr.execute('''DELETE FROM Employee WHERE Emp_id>480''')      #As there are 490 rows in table, all rows which are greater then 480 are deleted
    print(' Deletion successfull')
    conn.commit()
    conn.close()


def update_table():
    print('---------------------------------------------------------------------------------------------------------------------------------')
    print('Perform Updation for Table  Employee with Experience greater then 20 and with names Starting from S')
    conn = sqlite3.connect(ConStr)
    cr=conn.cursor()
    cr.execute('''UPDATE Employee
                SET Salary = 80000, Prev_Exp= Prev_Exp +1
                WHERE Emp_name LIKE 's%' and Prev_Exp>20;''')    #Updates the table by setting salary as 80000, and incrementing there Previous Experience
                                                          #for employee with name that starts and ends with s and a respectively and Experience>20
    print('Table Updated successfully')
    conn.commit()
    conn.close()

def select_cols_based_on_project():
    print('---------------------------------------------------------------------------------------------------------------------------------')
    print(' Query to get Employee\'s Id,Name, Age,Salary,Projects,Experience From table Employee based on their Project,Prev Experience and Salary') 
    print()
    conn = sqlite3.connect(ConStr)
    cr=conn.cursor()
    REC=cr.execute('''SELECT Emp_id,Emp_name,Salary,
    case
        when date(dob, '+' ||
            strftime('%Y', 'now') - strftime('%Y', dob) ||
            ' years') >= date('now')
        then strftime('%Y', 'now') - strftime('%Y', dob)
        else strftime('%Y', 'now') - strftime('%Y', dob) - 1
    end
    as Age,
	project,Prev_Exp
    FROM Employee WHERE PROJECT >150 and Prev_Exp >=15 and Salary>=80000  ORDER BY Project Desc;''') #Selects Emp_id,name,salary,age,project,Prev_Exp from Database
    print('Query Executed successfully')
    return REC
    conn.commit()
    conn.close()


def display_records(rec):
    print('---------------------------------------------------------------------------------------------------------------------------------')
    print('Display All Records from Query')
    conn = sqlite3.connect(ConStr)
    cr=conn.cursor()
    print()
    print('EMPLOYEE DETAILS BASED ON QUERY')
    print()
    print('Employee Id, '+'Employee Name, '+'Salary, '+'Age, '+'Project, '+'Prev Experience')  #Display Records stored in  database
    records=rec.fetchall()									#Fetch all query data and store in records variable
    for r in range(len(records)):
        print(str(records[r][0])+'         '+str(records[r][1])+'          '+str(records[r][2])+'     '+str(records[r][3])+'    '+str(records[r][4])+'       '+str(records[r][5]))

    return records
    conn.commit()
    conn.close()



def storing_newdata_excel(records,wb):
	
    print('---------------------------------------------------------------------------------------------------------------------------------')
    print('Storing Database Query details to Excel')
    wb.create_sheet(title='Updated Employee Data')
    sheet=wb.get_sheet_by_name('Updated Employee Data')

    sheet['A1'].font= Font(size=14, italic=True,bold=True)      #Styling Excel Sheets
    sheet['A1'] = 'EMPLOYEE\'S DATA FROM DATABASE'

    sheet['A2'],sheet['B2'],sheet['C2'],sheet['D2'],sheet['E2'],sheet['F2']='Emp Id ','Emp Name ','Salary ','Age ','Project ','Prev Experience'           #Set cell values

    i=0
    for rowNum in range(3,len(records)+3):				#Iterate to insert into Excel cells from 3rd row
        sheet.cell(row=rowNum,column=1).value,sheet.cell(row=rowNum,column=2).value,sheet.cell(row=rowNum,column=3).value=records[i][0],records[i][1],records[i][2] #Store record values
        sheet.cell(row=rowNum,column=4).value,sheet.cell(row=rowNum,column=5).value,sheet.cell(row=rowNum,column=6).value=records[i][3],records[i][4],records[i][5] #To excel sheet cells 
        i+=1
	
    print('  Data Stored Successfully ')

    wb.save('NewEmp_Data.xlsx')                          #Save WORKBOOK
    

def convert_xlsx_to_docx():
	
    print('---------------------------------------------------------------------------------------------------------------------------------')
    print(' Converting Excel to docx ')
    doc = docx.Document()                  #Create a Document

    workbook = op.load_workbook('NewEmp_Data.xlsx', data_only=True) #Load workbook NewEmp_Data which is result of query
    worksheet = workbook.get_sheet_by_name('Updated Employee Data')


    doc.add_paragraph('Employee Detail of Database query', 'Title').add_run()

    ws_range = worksheet.iter_cols(min_col=1, max_col=6, min_row=2, max_row=11, values_only=False)      #Produces cells from the worksheet, by column.
													#Specify the iteration range using indices of rows and columns.

    for row in ws_range:
        s = ''
        for cell in row:
            if cell.value is None:
                s += ' ' * 11

            else:
                s += str(cell.value).rjust(10) + ' '						       #Store data in s variable

        doc.add_paragraph(s)									       #Add stored data to word document as paragraph 
        doc.add_paragraph('________________________________________________________________________________________________________________________')
    doc.save('Excel_to_Doc.docx')




def convert_docx_to_pdf():
	
	print('---------------------------------------------------------------------------------------------------------------------------------')
	print('Converting Word Document to PDF')
    	convert('Excel_to_Doc.docx','C:/Users/srish/Desktop/PythonProject/Docx_to_Pdf.pdf')		#docx2pdf converts docs to pdf


def pdfmerge():
	
	print('---------------------------------------------------------------------------------------------------------------------------------')
	print('Merging Files')
	file1=open("Docx_to_Pdf.pdf", "rb")         #Opens pdf in read binary mode
	FirstFile = PyPDF2.PdfFileReader(file1)     #Reads all contents of infile1

	file2=open("MyProject.pdf", "rb")
	SecondFile = PyPDF2.PdfFileReader(file2)
	#watermark_pg=watermark.getPage(0).rotateClockwise(90)
	FinalFile = PyPDF2.PdfFileWriter()          #PdfFileWriter file holds the final file

	for i in range(FirstFile.getNumPages()):   #Iterates till highest Number of pages in PDF1
    		FinalFile.addPage(FirstFile.getPage(i))  #ADDS every page of firstfile to FinalFile

	for i in range(SecondFile.getNumPages()):   #Iterates till highest Number of pages in PDF2
    		FinalFile.addPage(SecondFile.getPage(i))  #ADDS every page of secondfile to FinalFile


	outfile = open("Final_PDF.pdf", 'wb')    #Open final pdf in write binary mode
	FinalFile.write(outfile)                #write each file

	file2.close()
	file1.close()                               # Closes all files that Was opened
	outfile.close()

def mail_automation():
	
	print('---------------------------------------------------------------------------------------------------------------------------------')
    	print('Initializing Mail Automation ')
    	browser=webdriver.Chrome()                                                      #Open Web Driver for Chrome
    	browser.get("http://gmail.com")                                                 #Get url to open gmail
    	email=browser.find_element_by_id('identifierId')

    	email.send_keys("SendersId@gmail.com")                                   #Sends senders mail id as key to browser
    	browser.find_element_by_id("identifierNext").click()                            #Clicks next button by finding element by id
    	time.sleep(3)                                                                   #waits for 3 seconds
    	browser.find_element_by_name("password").send_keys('Mypassword')                #Sends users password
    	browser.find_element_by_id("passwordNext").click()
    	time.sleep(3)                                                                   #waits for 3 seconds

    	browser.find_element_by_xpath('//*[@id=":iw"]/div/div').click()                 #Finds element by path
    	time.sleep(5)                                                                   #waits for 5 seconds
    	to=browser.find_element_by_name("to")                                           #finds element by name
    	to.send_keys("RecieversId@gmail.com")                                    #Send Recievers emailid
    	subject=browser.find_element_by_name("subjectbox")
    	subject.send_keys(" Python Automation Project")                                 #Sends Subject as key
    	MSG=browser.find_element_by_xpath('//*[@id=":ou"]')
    	MSG.send_keys('This mail is done with Automation using Selenium Webdriver.')                 #Sends Body Messages
    	MSG.send_keys(' Please do find our Github repo link below with python code and all  files that were genetrated while running  ')
    	MSG.send_keys('https://github.com/Srish283/PythonAutomationProject')
    	time.sleep(2)
    	attach=browser.find_element_by_xpath('//*[@id=":nh"]').click()
    	time.sleep(5)                                                                   #waits for 5 seconds
    	browser.close()                                                                 #Close browser




if __name__=="__main__":
	
    print('Mini Project to Demonstrate Working of following Concepts:  ')
    print('---------------------------------------------------------------------------------------------------------------------------------')
   
    print('1. Excel: Reading and Writing To Excel files with openpyxl module')
    print('2. Database: Perform CRUD operation by connecting to Sqlite3')
    print('3. Word: Convert Excel file of with query results to Word with pydocx module')
    print('4. PDF: Convert Word to Pdf and Merge different PDF files with PyPDF2 module')
    print('5. Selenium: Automate mail with Selenium webdriver')
	
    print('---------------------------------------------------------------------------------------------------------------------------------')
    print(' Let\'s  Get Started ')
    print()
    ConStr="EmpExcelData.db"
								
    sheet,wb=open_workbook()				# Function Calls
    create_table()
    insert_into_table(sheet)
    delete_rows()
    update_table()
    REC=select_cols_based_on_project()
    records=display_records(REC)
    storing_newdata_excel(records,wb)
    convert_xlsx_to_docx()
    convert_docx_to_pdf()
    pdfmerge()
    mail_automation()
    print(' Thank You  :)')

