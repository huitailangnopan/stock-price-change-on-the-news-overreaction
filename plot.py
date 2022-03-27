import requests
import os
import openpyxl
import time
import json
from match import match
import matplotlib
import pandas as pd
import matplotlib.pyplot as plt

os.chdir(r"D:\high-frequency data\pharmarcy")
arr = os.listdir()
stockfolder=[]
for file in arr:
   if (file.find('.xlsx') == -1):
       if (file.find('.py') == -1):
            stockfolder.append(file)
for i in range(0,len(stockfolder)):
    ticker = stockfolder[i]
    os.chdir("D:\\high-frequency data\\analysis")
    df = pd.read_excel(ticker+'.xlsx')
    dk = pd.DataFrame(columns=['one_min','one_hour','pre_increase','daily_increase','post_increase','next_day','next_10days','next_month'])
    dk.at[0, 'pre_increase'] = df['pre_increase'].mean()
    dk.at[0, 'daily_increase'] = df['daily_increase'].mean()
    dk.at[0, 'post_increase'] = df['post_increase'].mean()
    dk.at[0, 'one_min'] = df['one_min'].mean()
    dk.at[0, 'one_hour'] = df['one_hour'].mean()
    dk.at[0, 'next_day'] = df['next_day'].mean()
    dk.at[0, 'next_10days'] = df['next_10days'].mean()
    dk.at[0, 'next_month'] = df['next_month'].mean()
    os.chdir("D:\\high-frequency data\\analysis\\plt")
    #plt.plot(dk.index, dk).savefig(ticker+'.png')
    #plt.hist(dk)
    #plt.savefig(ticker+'.png')
    ax = dk.plot.bar()
    ax.figure.savefig(ticker+'.png')
