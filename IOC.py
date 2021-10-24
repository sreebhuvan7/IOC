"""
@author: G Sree Bhuvan
         PhD Student, Department of Earth Sciences
         Pondicherry University
"""

import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np
import os
from IPython.display import display

root1= tk.Tk()
root1.title('Iron Oxide Converter')

canvas2 = tk.Canvas(root1, width = 300, height = 300, bg = 'white')
canvas2.pack()

def getExcel():
    global df

    import_file_path = filedialog.askopenfilename()

    df = pd.read_excel(import_file_path)

    root1.destroy()

browseButton_Excel1 = tk.Button(text='Import Excel File', command=getExcel, bg='black', fg='forest green', font=('oxygen', 12, 'bold'))
canvas2.create_window(150, 150, window=browseButton_Excel1)

root1.lift()
root1.attributes('-topmost',True)
root1.mainloop()

l = 0
m = 0
x = []
for m in df["SiO2"]:

    if (df.loc[l,"SiO2"]<=52) and (df.loc[l,"SiO2"]>=44):
        x.append(0.2)
        l+=1
    elif (df.loc[l,"SiO2"]>=52) and (df.loc[l,"SiO2"]<=57):
        x.append(0.3)
        l+=1
    elif (df.loc[l,"SiO2"]>=57) and (df.loc[l,"SiO2"]<=62):
        x.append(0.35)
        l+=1
    elif (df.loc[l,"SiO2"]>=62) and (df.loc[l,"SiO2"]<=77):
        x.append(0.4)
        l+=1
    elif (df.loc[l,"SiO2"]>=77):
        x.append(0.5)
        l+=1
    elif (df.loc[l,"MgO"]>18):
        x.append(0.12)
        l+=1
m+=1
df.append(x)

df['x'] = x

a = 0
g = 0.8998
h = 1.1113

if 'FeOt' in df:
    j = df['FeOt']
    for y in df['x']:
        d = df.loc[a:,'x']
        FeO = round(j/(1+(d*g)),2)
        Fe2O3 = round(d*FeO,2)
        Fe2O3t = round((h*FeO)+Fe2O3,2)
    df['FeO'] = FeO
    df['Fe2O3'] = Fe2O3
    df['Fe2O3t'] = Fe2O3t
    y+=1
    a+=1

elif 'FeOT' in df:
    j = df['FeOT']
    for y in df['x']:
        d = df.loc[a:,'x']
        FeO = round(j/(1+(d*g)),2)
        Fe2O3 = round(d*FeO,2)
        Fe2O3t = round((h*FeO)+Fe2O3,2)
    df['FeO'] = FeO
    df['Fe2O3'] = Fe2O3
    df['Fe2O3t'] = Fe2O3t
    y+=1
    a+=1

elif 'FeO(t)' in df:
    j = df['FeO(t)']
    for y in df['x']:
        d = df.loc[a:,'x']
        FeO = round(j/(1+(d*g)),2)
        Fe2O3 = round(d*FeO,2)
        Fe2O3t = round((h*FeO)+Fe2O3,2)
    df['FeO'] = FeO
    df['Fe2O3'] = Fe2O3
    df['Fe2O3t'] = Fe2O3t
    y+=1
    a+=1
elif 'FeO(T)' in df:
    j = df['FeO(T)']
    for y in df['x']:
        d = df.loc[a:,'x']
        FeO = round(j/(1+(d*g)),2)
        Fe2O3 = round(d*FeO,2)
        Fe2O3t = round((h*FeO)+Fe2O3,2)
    df['FeO'] = FeO
    df['Fe2O3'] = Fe2O3
    df['Fe2O3t'] = Fe2O3t
    y+=1
    a+=1

elif 'Fe2O3t' in df:
    j = df['Fe2O3t']
    for y in df['x']:
        d = df.loc[a:,'x']
        FeO = round(j/(d+h),2)
        Fe2O3 = round(d*FeO,2)
        FeOt = round(FeO+(g*Fe2O3),2)
    df["FeO"] = FeO
    df["Fe2O3"] = Fe2O3
    df["FeOt"] = FeOt
    y+=1
    a+=1

elif 'Fe2O3(t)' in df:
    j = df["Fe2O3(t)"]
    for y in df['x']:
        d = df.loc[a:,'x']
        FeO = round(j/(d+h),2)
        Fe2O3 = round(d*FeO,2)
        FeOt = round(FeO+(g*Fe2O3),2)
    df["FeO"] = FeO
    df["Fe2O3"] = Fe2O3
    df["FeOt"] = FeOt
    y+=1
    a+=1

elif 'Fe2O3T' in df:
    j = df["Fe2O3T"]
    for y in df['x']:
        d = df.loc[a:,'x']
        FeO = round(j/(d+h),2)
        Fe2O3 = round(d*FeO,2)
        FeOt = round(FeO+(g*Fe2O3),2)
    df["FeO"] = FeO
    df["Fe2O3"] = Fe2O3
    df["FeOt"] = FeOt
    y+=1
    a+=1

elif "Fe2O3(T)" in df:
    j = df["Fe2O3(T)"]
    for y in df['x']:
        d = df.loc[a:,'x']
        FeO = round(j/(d+h),2)
        Fe2O3 = round(d*FeO,2)
        FeOt = round(FeO+(g*Fe2O3),2)
    df["FeO"] = FeO
    df["Fe2O3"] = Fe2O3
    df["FeOt"] = FeOt
    y+=1
    a+=1

elif 'FeO' and 'Fe2O3' in df:
    FeO = df['FeO']
    Fe2O3 = df['Fe2O3']
    Fe2O3t = round(Fe2O3+(h*FeO),2)
    FeOt = round((g*Fe2O3)+FeO,2)
    df["Fe2O3t"] = Fe2O3t
    df["FeOt"] = FeOt

elif 'FeO' in df:
    FeO = df['FeO']
    Fe2O3 = round(d*FeO,2)
    Fe2O3t = round(Fe2O3+(h*FeO),2)
    FeOt = round((g*Fe2O3)+FeO,2)
    df["Fe2O3"] = Fe2O3
    df["Fe2O3t"] = Fe2O3t
    df["FeOt"] = FeOt

elif 'Fe2O3' in df:
    Fe2O3 = df['Fe2O3']
    FeO = round(x*Fe2O3,2)
    Fe2O3t = round(Fe2O3+(h*FeO),2)
    FeOt = round((g*Fe2O3)+FeO,2)
    df["FeO"] = Fe2O3
    df["Fe2O3t"] = Fe2O3t
    df["FeOt"] = FeOt


else:
        print('Check your input!')


writer = pd.ExcelWriter("Iron Oxide Converted.xlsx", engine='xlsxwriter')
df.to_excel(writer, sheet_name = 'Sheet 1', index=False)
writer.save()
