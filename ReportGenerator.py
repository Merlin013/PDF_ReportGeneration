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
    drive = '(Logical Disk:)'
    s_name = []
    drive_name = []
    for i in range(0, num_pages):
        page = pdfRead.getPage(i)
        text += page.extractText()
        text = text.strip()
    try:
        for line in io.StringIO(text):
            if re.findall(drive, line):
                # print(line)
                pos = line.find(":")
                # words1 = line.split()
                # print(words1)
                # if re.search('(\d+)',words1[pos+1]):
                #     pass

                drive_name.append(line[pos+1:-1])
        print(drive_name)

        docu = open("Drive names.txt", "a+")
        # doc.write("List of Drive names\n")
        for i in range(len(drive_name)):
            docu.write(drive_name[i] + "\n")

        docu.close()
    except:
        pass

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


# def drivenames():
#     dfmain.insert(1, "DriveName", " ")
#     print(dfmain)
#
#     dfr = dfmain.loc[dfmain["Interval"] == "Interval"]
#
#     print(dfr)
#     rowval = list(dfr.index)
#     rowval.append(0)
#     rowval.sort()
#     print(rowval)
#     doc = open("Drive names.txt", "r")
#     lines = doc.readlines()
#     dname = []
#     count = 1
#     for j in lines:
#         dname.append(j)
#     for x in range(len(rowval)):
#         try:
#             while (count < rowval[x + 1]):
#                 dfmain.at[count, "DriveName"] = dname[x]
#                 count += 1
#         except:
#             if count == rowval[-1]:
#                 while (count <= count_rows):
#                     dfmain.at[count, "DriveName"] = dname[-1]
#                     count += 1


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
    dfmain.insert(1, "DriveName", " ")
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
    docu = open("Drive names.txt", "r")
    lines = doc.readlines()
    lines1 = docu.readlines()
    sname = []
    dname = []
    count = 0 # changed to 0 make 1 later
    for j in lines:
        sname.append(j)

    for j in lines1:
        dname.append(j)

    for x in range(len(rowval)):
        try:
            while (count < rowval[x+1]):
                dfmain.at[count, "ServerName"] = sname[x]
                dfmain.at[count, "DriveName"] = dname[x]
                count+=1

        except:
            if count == rowval[-1]:
                while (count < count_rows):
                    dfmain.at[count, "ServerName"] = sname[-1]
                    dfmain.at[count, "DriveName"] = dname[-1]
                    count += 1

    # try:
    #     if os.path.exists("Drive names"):
    #         drivenames()
    #
    # except:
    #     pass

    rname = input("Enter final report name with .xlsx extension\n")
    writer = pf.ExcelWriter(rname, engine='openpyxl', mode='w,a')
    dfmain.to_excel(writer, index=False)
    writer.save()
    doc.close()
    docu.close()

    end = time.time()
    print(f"Total time taken:{(end-begin)/60} minutes")
    time.sleep(10)
    os.remove(output_csv)
    os.remove("Server names.txt")
    os.remove("Drive names.txt")
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