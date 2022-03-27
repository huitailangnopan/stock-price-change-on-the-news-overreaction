import requests
import os
import openpyxl
import time
import json
from match import match
'''
os.chdir(r"D:\high-frequency data\pharmarcy")
arr = os.listdir()
stockfolder=[]
for file in arr:
   if (file.find('.xlsx') == -1):
       if (file.find('.py') == -1):
            stockfolder.append(file)
for q in range(len(stockfolder)):

'''


os.chdir(r"D:\high-frequency data\pharmarcy")
arr = os.listdir()
stockfolder=[]
for file in arr:
   if (file.find('.xlsx') == -1):
       if (file.find('.py') == -1):
            stockfolder.append(file)
for i in range(len(stockfolder)):
    k = match(stockfolder[i])
