"""This program keeps track of Asylum Farming in the game Fallout 76
    It uses Tkinter for a UI that allows user input, appends inputs to an array,
    then ultimately adds those arrays to an excel spreadsheet using PANDAS """


import tkinter as tk
import time
from pynput import keyboard
import pandas as pd
import numpy as np
import openpyxl as xl
from datetime import datetime
import functools
import os

runNumber = 1
data = np.array([int(),int(),int(),int(),int(),int(),int(),int(),int(),int(),int()])
now = datetime.now()
iterationNumber = 1


def btnLogStart_clicked():
    global runNumber
    global data
    global iterationNumber
    now = datetime.now()
    current_time = now.strftime("%Y-%m-%d %H:%M:%S")
    if iterationNumber == 1:
        newrow = [runNumber,iterationNumber,current_time,0,0,0,0,0,0,0,0]
        data = np.vstack([newrow])
    print (data)

def btnLogIteration_clicked():
    global runNumber
    global data
    global iterationNumber
    now = datetime.now()
    current_time = now.strftime("%Y-%m-%d %H:%M:%S")
    iterationNumber +=1
    newrow = [runNumber,iterationNumber,current_time,Brown_Spin_Val.get(),Green_Spin_Val.get(),Blue_Spin_Val.get(),Pink_Spin_Val.get(),Yellow_Spin_Val.get(),Forest_Spin_Val.get(),Red_Spin_Val.get(),Total_Spin_Val.get()]
    data = np.vstack([data, newrow])
    print (data)

def btnLogRun_clicked():
    global runNumber
    global data
    global iterationNumber
    now = datetime.now()
    current_time = now.strftime("%Y-%m-%d %H:%M:%S")
    iterationNumber =1
    runNumber +=1
    Brown_Spin_Val.set(0)
    Green_Spin_Val.set(0)
    Blue_Spin_Val.set(0)
    Pink_Spin_Val.set(0)
    Yellow_Spin_Val.set(0)
    Forest_Spin_Val.set(0)
    Red_Spin_Val.set(0)
    newrow = [runNumber,iterationNumber,current_time,Brown_Spin_Val.get(),Green_Spin_Val.get(),Blue_Spin_Val.get(),Pink_Spin_Val.get(),Yellow_Spin_Val.get(),Forest_Spin_Val.get(),Red_Spin_Val.get(),Total_Spin_Val.get()]
    data = np.vstack([data, newrow])
    print (data)

def btnLogEnd_clicked():
    global runNumber
    global data
    global iterationNumber
    now = datetime.now()
    current_time = now.strftime("%Y-%m-%d %H:%M:%S")
    newrow = [runNumber,iterationNumber,current_time,0,0,0,0,0,0,0,0]
    data = np.vstack([data, newrow])
    #newrow = ['Total Attempts','Total Time','Total Brown','Total Green','Total Blue','Total Yellow','Total Forrest','Total Red','Total Spawns']
    #data = np.vstack([data, newrow])
    #newrow = [CountCounter-1,'=SUM(C2:C'+str(CountCounter+2)+')','=SUM(C2:C'+str(CountCounter+2)+')','=SUM(D2:D'+str(CountCounter+2)+')','=SUM(E2:E'+str(CountCounter+2)+')','=SUM(F2:F'+str(CountCounter+2)+')','=SUM(G2:G'+str(CountCounter+2)+')','=SUM(H2:H'+str(CountCounter+2)+')','=SUM(I2:I'+str(CountCounter+2)+')']
    #data = np.vstack([data, newrow])
    print (data)
    df = pd.DataFrame(data)
    df.apply(pd.to_numeric,errors='ignore').info()
    #df.iloc[1:] = df.astype(int, columns=['Run','Iteration','Time','Brown','Green', 'Blue','Pink','Yellow','Forrest','Red','Total'])
    #df['Run','Iteration','Brown','Green', 'Blue','Pink','Yellow','Forrest','Red','Total'] = df[['Run','Iteration','Brown','Green', 'Blue','Pink','Yellow','Forrest','Red','Total']].astype(int)
    if os.path.exists("C:/Users/msujo/Desktop/Projects/Fallout Asylum/Data Output/%s.xlsx" %(datetime.today().strftime('%Y-%m-%d'))) == False:
        df.to_excel("C:/Users/msujo/Desktop/Projects/Fallout Asylum/Data Output/%s.xlsx" %(datetime.today().strftime('%Y-%m-%d')),sheet_name="sheet",index = None, header = ['Run','Iteration','Time','Brown','Green', 'Blue','Pink','Yellow','Forrest','Red','Total'])
    else:
        ExcelFile = "C:/Users/msujo/Desktop/Projects/Fallout Asylum/Data Output/%s.xlsx" %(datetime.today().strftime('%Y-%m-%d'))
        WorkBook = xl.load_workbook(ExcelFile)
        res = len(WorkBook.sheetnames)                                                           
        count=res+1
        count2 = count 
        with pd.ExcelWriter("C:/Users/msujo/Desktop/Projects/Fallout Asylum/Data Output/%s.xlsx" %(datetime.today().strftime('%Y-%m-%d')), engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name="sheet"+str(count),index = None, header = ['Run','Iteration','Time','Brown','Green', 'Blue','Pink','Yellow','Forrest','Red','Total'])
            df['Run','Iteration','Brown','Green', 'Blue','Pink','Yellow','Forrest','Red','Total'] = df ['Run','Iteration','Time','Brown','Green', 'Blue','Pink','Yellow','Forrest','Red','Total'].astype(int)
    window.destroy()
    

window = tk.Tk()

window.title('FO76 Data Collection Program')

window.tk.call('tk', 'scaling', 2.5)
fontBold=("Serif",12,'bold')

header = tk.Label(window, text= 'This program logs and exports data on \n Asylum Dress farming in the game Fallout 76', font =fontBold)
header.grid(column=0, row=0,columnspan=2)

RunCount = tk.Label(window, text= 'Number of Runs this session (Run = series of interations):')
RunCount.grid(column=0, row=2)


BrownCount = tk.Label(window, text= 'Number of Brown Dresses this iteration')
BrownCount.grid(column=0, row=3)
Brown_Spin_Val = tk.IntVar()
BrownSpin = tk.Spinbox(window, from_=0, to=3, textvariable=Brown_Spin_Val)
BrownSpin.grid(column=1, row=3)

GreenCount = tk.Label(window, text= 'Number of Green Dresses this iteration')
GreenCount.grid(column=0, row=4)
Green_Spin_Val = tk.IntVar()
GreenSpin = tk.Spinbox(window, from_=0, to=3, textvariable=Green_Spin_Val)
GreenSpin.grid(column=1, row=4)

BlueCount = tk.Label(window, text= 'Number of Blue Dresses this iteration')
BlueCount.grid(column=0, row=5)
Blue_Spin_Val = tk.IntVar()
BlueSpin = tk.Spinbox(window, from_=0, to=3, textvariable=Blue_Spin_Val)
BlueSpin.grid(column=1, row=5)

PinkCount = tk.Label(window, text= 'Number of Pink Dresses this iteration')
PinkCount.grid(column=0, row=6)
Pink_Spin_Val = tk.IntVar()
PinkSpin = tk.Spinbox(window, from_=0, to=3, textvariable=Pink_Spin_Val)
PinkSpin.grid(column=1, row=6)

YellowCount = tk.Label(window, text= 'Number of Yellow Dresses this iteration')
YellowCount.grid(column=0, row=7)
Yellow_Spin_Val = tk.IntVar()
YellowSpin = tk.Spinbox(window, from_=0, to=3, textvariable=Yellow_Spin_Val)
YellowSpin.grid(column=1, row=7)

ForestCount = tk.Label(window, text= 'Number of Forest Dresses this iteration')
ForestCount.grid(column=0, row=8)
Forest_Spin_Val = tk.IntVar()
ForestSpin = tk.Spinbox(window, from_=0, to=3, textvariable=Forest_Spin_Val)
ForestSpin.grid(column=1, row=8)

RedCount = tk.Label(window, text= 'Number of Red Dresses this iteration')
RedCount.grid(column=0, row=9)
Red_Spin_Val = tk.IntVar()
RedSpin = tk.Spinbox(window, from_=0, to=3, textvariable=Red_Spin_Val)
RedSpin.grid(column=1, row=9,)

TotalCount = tk.Label(window, text= 'Number of Total Dresses this iteration')
TotalCount.grid(column=0, row=10)
Total_Spin_Val = tk.IntVar()
TotalSpin = tk.Spinbox(window, from_=0, to=3, textvariable=Total_Spin_Val)
TotalSpin.grid(column=1, row=10,)

btnLogStart = tk.Button(window, text="Begin this Session", command=btnLogStart_clicked)
btnLogStart.grid(column=0, row=11, columnspan=2)

btnLogIteration = tk.Button(window, text="Log Iteration", command=btnLogIteration_clicked)
btnLogIteration.grid(column=0, row=12, columnspan=2)

btnLogRun = tk.Button(window, text="Log Run", command=btnLogRun_clicked)
btnLogRun.grid(column=0, row=13, columnspan=2)

btnLogEnd = tk.Button(window, text="End and Export", command=btnLogEnd_clicked)
btnLogEnd.grid(column=0, row=14, columnspan=2)



window.mainloop()


