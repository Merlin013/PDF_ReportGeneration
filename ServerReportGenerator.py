import PyPDF2 as py
import re
import io
import os
import tabula
from tabulate import tabulate
import xlsxwriter
import openpyxl
import pandas as pf
import time

'''
This code was created by Vishal Petkar on 16th Oct 2020. 
Github - merlin013
'''

def sdname(pfile):
    pdfRead = py.PdfFileReader(pfile)
    num_pages = pdfRead.numPages
    text = ""
    regex = '(SCOM)'
    # drive = '(Logical Disk:)'
    s_name = []
    for i in range(0, num_pages):
        page = pdfRead.getPage(i)
        text += page.extractText()
        text = text.strip()


    for line in io.StringIO(text):
        if re.findall(regex, line):
            words = line.split()
            server = words[2]
            sname = server.split('.')
            s_name.append(sname[0])
            print(sname[0])

    doc = open("Server names.txt", "a+")
    # doc.write("List of server names\n")
    for i in range(len(s_name)):
        doc.write(s_name[i] + "\n")

    doc.close()


if __name__ == '__main__':
    begin = time.time()
    fname = input("Enter file name\n")
    pfile = open(fname, "rb")
    sdname(pfile)
    output_csv = "Buffer.csv"
    tabula.convert_into(fname, output_path=output_csv, output_format="csv", stream=True, pages='all')

    df = pf.read_csv(output_csv)
    df1 = df[df.columns[0:6]]
    count_rows = df1.shape[0]
    dfmain = df1.drop(columns= ["Sample Count", "Standard Deviation"])
    # dfmain['ServerName'] = None
    dfmain.insert(0, "ServerName", " " )
    print(dfmain)
    # rowval = []
    # for row in range(count_rows):
    dfr = dfmain.loc[dfmain["Interval"] == "Interval"]
    # rowval.append(row)
    print(dfr)
    rowval = list(dfr.index)
    rowval.append(0)
    rowval.sort()
    print(rowval)
    doc = open("Server names.txt", "r")
    lines = doc.readlines()
    sname = []
    count = 0
    for j in lines:
        sname.append(j)
    for x in range(len(rowval)):
        try:
            while (count < rowval[x+1]):
                dfmain.at[count, "ServerName"] = sname[x]
                count+=1
        except:
            if count == rowval[-1]:
                while (count < count_rows):
                    dfmain.at[count, "ServerName"] = sname[-1]
                    count+=1

    rname = input("Enter final report name with .xlsx extension\n")
    writer = pf.ExcelWriter(rname, engine='openpyxl', mode='w,a')
    dfmain.to_excel(writer, index=False)
    writer.save()
    doc.close()
    end = time.time()
    print(f"Total time taken:{(end - begin) / 60} minutes")
    time.sleep(10)
    os.remove(output_csv)
    os.remove("Server names.txt")