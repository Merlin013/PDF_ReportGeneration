import ReportGenerator
import ServerReportGenerator
import os

'''
This code was created by Vishal Petkar on 16th Oct 2020. 
Github - merlin013
'''

choice= int(input("*****Menu*****\n1.Server Details\n2.Disk/Drive Details\nSelect 1 for generating "
                  "Server details file or 2 for generating a Disk/drive file\n-->"))

if choice == 1:
    os.system('python ServerReportGenerator.py')
elif choice == 2:
    os.system('python ReportGenerator.py')
else:
    print("Please try again")