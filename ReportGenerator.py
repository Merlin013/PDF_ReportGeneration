import PyPDF2 as py
import re
import io
# import tkinter
# import ghostscript
# import camelot as cam
from tabula import read_pdf
from tabulate import tabulate
import xlsxwriter
import openpyxl
import pandas as pf

fname = input("Enter file name\n")
pfile = open(fname, "rb")
pdfRead = py.PdfFileReader(pfile)
num_pages = pdfRead.numPages
text = ""
regex = '(SCOM)'
drive = '(Logical Disk:)'

for i in range(0, num_pages):
    page= pdfRead.getPage(i)
    text += page.extractText()
    text = text.strip()
try:
    for line in io.StringIO(text):
        if re.findall(drive, line):
            print(line)
            words1 = line.split()
            print(words1)
            drive_name = words1[7]
            print(drive_name)
except:
    pass

for line in io.StringIO(text):
    if re.findall(regex, line):
        words = line.split()
        server = words[2]
        sname = server.split('.')
        print(sname[0])



# Camelot does not work with this PDF
# tables = cam.read_pdf(fname)
# print("Total tables extracted:", tables.n)
#
# print(tables[0].df)

df = read_pdf(fname, pages = 'all', multiple_tables= True, encoding = 'ISO-8859-1', stream=True)
# print(tabulate(df))

# writer = pf.ExcelWriter('SLAreport.xlsx', engine='openpyxl', mode='w,a')
df1 = pf.DataFrame(df)
# df1.to_excel(writer)
print(tabulate(df1))
#
# writer.save()