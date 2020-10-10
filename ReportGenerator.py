import PyPDF2 as py
import re
import io
# import tkinter
# import ghostscript
# import camelot as cam
import tabula
from tabulate import tabulate
import xlsxwriter
import openpyxl
import pandas as pf

def sdname(pfile):
    pdfRead = py.PdfFileReader(pfile)
    num_pages = pdfRead.numPages
    text = ""
    regex = '(SCOM)'
    drive = '(Logical Disk:)'
    s_name = []
    for i in range(0, num_pages):
        page = pdfRead.getPage(i)
        text += page.extractText()
        text = text.strip()
    try:
        for line in io.StringIO(text):
            if re.findall(drive, line):
                print(line)
                words1 = line.split()
                print(words1)
                if re.search('(\d+)',words1[7]):
                    pass
                drive_name = words1[7]
                print(drive_name)
    except:
        pass

    for line in io.StringIO(text):
        if re.findall(regex, line):
            words = line.split()
            server = words[2]
            sname = server.split('.')
            s_name.append(sname[0])
            print(sname[0])

    doc = open("Server names", "a+")
    # doc.write("List of server names\n")
    for i in range(len(s_name)):
        doc.write(s_name[i] + "\n")

    doc.close()


if __name__ == '__main__':

    fname = input("Enter file name\n")
    pfile = open(fname, "rb")
    sdname(pfile)
    output_csv = "Buffer.csv"
    tabula.convert_into(fname, output_path=output_csv, output_format="csv", stream=True, pages='all')

    df = pf.read_csv(output_csv)
    df1 = df[df.columns[0:6]]
    count_rows = df1.shape[0]
    dfmain = df1.drop(columns= ["Sample Count", "Min Value", "Standard Deviation"])
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
    doc = open("Server names", "r")
    lines = doc.readlines()
    sname = []
    for j in lines:
        sname.append(j)
    for x in range(len(rowval)):
        dfmain.at[rowval[x]+1,"ServerName"] = sname[x]

    writer = pf.ExcelWriter('SLAreport.xlsx', engine='openpyxl', mode='w,a')
    dfmain.to_excel(writer, index=False)
    writer.save()

# global df1, df

# Camelot does not work with this PDF
# tables = cam.read_pdf(fname)
# print("Total tables extracted:", tables.n)
#
# print(tables[0].df)
# for i in range(1, num_pages):
# df = tabula.read_pdf(fname, pages = 'all', multiple_tables= True, encoding='utf-8', stream = True, output_format='dataframe')
# # df1 = pf.DataFrame(columns=['Index', 'Interval', 'Sample Count', 'Min Value', 'Max Value', 'Average Value', 'Standard Deviation'])
#
# # headers = df.pop(0)
# # print(headers)
# print(len(df))
# print(df)
# print(type(df))
# df1 = pf.DataFrame(df, columns= ['Index', 'Interval', 'Sample Count', 'Min Value', 'Max Value', 'Average Value', 'Standard Deviation'])
# print(type(df1))
# print(df1)
# print(tabulate(df, showindex=True))

# df.to_csv("output2.csv", encoding='utf-8')
# df1 = pf.DataFrame(columns=['Index', 'Interval', 'Sample Count', 'Min Value', 'Max Value', 'Average Value',
#                                     'Standard Deviation']).from_dict(df)
#
#
#
#
# # df1 = pf.DataFrame(df, columns=['Index', 'Interval', 'Sample Count', 'Min Value', 'Max Value', 'Average Value', 'Standard Deviation'])
# df1.to_excel(writer, startcol=1)
# # print(tabulate(df1, showindex=False))
# #
# writer.save()