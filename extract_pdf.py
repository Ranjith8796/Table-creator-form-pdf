import os
import sys
import csv
import pdfplumber
from PyPDF2 import PdfFileReader,PdfFileWriter
from reportlab.pdfgen import canvas

##################################################################################################################

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(__file__)

file=open(os.path.join(BASE_DIR,'100100502_AES01267KY_ORG_Financial-Statement_2020.pdf'),'rb')

##################################################################################################################

def create_separate_page():
    base_file_name='page'
    for n in range(3,5):
        existing_pdf = PdfFileReader(file,'rb')
        output = PdfFileWriter()
        output.addPage(existing_pdf.getPage(n))   
        with open(os.path.join(BASE_DIR,'{}_subset_{}.pdf'.format(base_file_name,n)),'wb') as f:
            output.write(f)
            f.close

def create_pdf_line():
    for number in range(3,5):
        if number==3:
            realpdf3=canvas.Canvas(os.path.join(BASE_DIR,'real_pdf_3.pdf'))
            realpdf3.line(40,168,40,655)
            realpdf3.line(561.5,168,561.5,655)
            realpdf3.save()
        else:
            realpdf4=canvas.Canvas(os.path.join(BASE_DIR,'real_pdf_4.pdf'))
            realpdf4.line(51,178,51,655)
            realpdf4.line(573,178,573,655)
            realpdf4.save()

def merging_two_pdf():
    for num in range(3,5):
        line=PdfFileReader(os.path.join(BASE_DIR,'real_pdf_'+str(num)+'.pdf'))
        table=PdfFileReader(os.path.join(BASE_DIR,'page_subset_'+str(num)+'.pdf'))
        outputpage = PdfFileWriter()
        page = table.getPage(0)
        page.mergePage(line.getPage(0))
        outputpage.addPage(page)
        with open(os.path.join(BASE_DIR,'final_'+str(num)+'.pdf'),'wb') as outputStream:
            outputpage.write(outputStream)
            outputStream.close

def create_csv_file(no):
    csvdata=[]
    pdf = pdfplumber.open(os.path.join(BASE_DIR,'final_'+str(no)+'.pdf')) 
    os.remove(os.path.join(BASE_DIR,'real_pdf_'+str(no)+'.pdf'))
    os.remove(os.path.join(BASE_DIR,'page_subset_'+str(no)+'.pdf'))
    for page in pdf.pages[:1]:                    
        for table in page.extract_tables():
            for row in table:
                csvdata.append(row)
    pdf.close()
    return csvdata

##################################################################################################################

create_pdf_line()
create_separate_page()
merging_two_pdf()


for no in range(3,5):
    data=create_csv_file(no)
    csv_flile=open(os.path.join(BASE_DIR,'Table_of_pdf_page_'+str(no+1)+'.csv'), 'w+', newline ='')
    with csv_flile:    
        write = csv.writer(csv_flile)
        write.writerows(data)
    os.remove(os.path.join(BASE_DIR,'final_'+str(no)+'.pdf'))







