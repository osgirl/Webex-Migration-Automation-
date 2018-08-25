# -*- coding: utf-8 -*-
"""
Created on Fri May 18 08:18:19 2018

@author: zachary.shaver
"""

import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import os 
from pathlib import Path
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")

def send_mail(address, body):
    mail = outlook.CreateItem(0)
    mail.To = address
    mail.Subject = 'SCHEDULE INFO'
    mail.Body = body
    #print('---------------------------------------------------------------------------', address, body)
    mail.Send()
    
def retard_corrector(name):
    if name == 'Kendra':
        return 'Khendra'
    elif name == 'Donte':
        return 'Dontavius'
    else:
        return name

master_list = []

email_list = {'Khendra': "khendra.davidson@baesystems.com",
              'Dylan': 'dylan.drum@baesystems.com', 
              'Dontavius': 'dontavius.hopkins@baesystems.com', 
              'Jaylan': 'jaylan.mobley@baesystems.com', 
              'Yuliya': 'yuliya.pinchuk@baesystems.com', 
              'Zachary': 'zachary.shaver@baesystems.com', 
              'Sean': 'sean.stewart2@baesystems.com', 
              'Alina': 'alina.svarishchuk@baesystems.com', 
              'George': 'george.vargas@baesystems.com', 
              'Darnell': 'darnell.wallace@baesystems.com',
              'Brandon': 'brandon.wells@baesystems.com'}



catcher = True
in_days = [] 

c = input("Enter the month (Exact month case sensitive from the FedRAMP Migration 2018/Sign Up Lists/ - ")
folder_path = Path("//wtooleast.nosc.na.baesystems.com/Network_Documentation/Enterprise_Documentation/Voice/FedRAMP Migration 2018/Sign Up  Lists" + "/" + c)
print(folder_path.exists())


for file in os.listdir(folder_path):
    
    
    filename = os.fsdecode(file)
    
    #create the day for each day of the loop 
    day_master = {"day": "", 'Khendra': [], 'Dylan': [], 'Dontavius': [],'Jaylan': [],'Yuliya': [],'Zachary': [], 'Sean': [], 'Alina': [], 'George': [], 'Darnell': [], 'Brandon': []}

    print("are the lists for ", filename, "complete y/n")
    check = input("y/n (case sensitive)")
    if "x" in check:
        break
    if "y" not in check:
        continue
    
    day = filename.split(".")[0]
    day_master['day'] = day
    #print("DATE  ---", filename.split('.')[0])
    
    #go through directory for each time in the day
    for subfile in os.listdir(str(folder_path)+'/'+filename):
        
        f_nam = os.fsdecode(subfile)
        time = f_nam.split('.')[0]
        if f_nam.split('.')[1] == 'xlsx':
            
            if "-" not in f_nam or " " in f_nam:
                continue
    
            #print("TIME ----", os.fsdecode(subfile))
            
            #dict that holds the emails and times of each trainer             
            
            #excell shit               
            ws = load_workbook(str(folder_path)+'/'+filename+'/'+str(subfile))
            timesheet = ws.get_sheet_by_name(ws.get_sheet_names()[0])
            
            #email row is always the same trainers is not for some fucking reason 
            emails = timesheet['B']
            
            if 'Trainer' in timesheet['E'][15].value:
                trainers = timesheet['E']
            elif 'Attendance' in timesheet['E'][15].value:
                trainers = timesheet['D']
                
            
            prev = trainers[16].value
            
            
            hour_master = os.fsdecode(subfile).split('.')[0]
            email_master = []
            #time_master = 0
            
            #go through each excell doc to find the times for the specific person 
            #print(len(emails))
            for x in range(16, len(emails)):
                #print("Email - ", emails[x].value, "Person - ", trainers[x].value)
                if prev is None or emails[x].value is None:
                    continue
                if prev == trainers[x].value:
                    email_master.append(emails[x].value.split(" ")[0])
                    #time_master += hours[x].value
                else:
                    prev = retard_corrector(prev)
                    
                    #print("ran", emails[x].value, trainers[x].value, hour_master)
                    day_master[prev.split(" ")[0]].append([hour_master, email_master])
                    email_master = []
                    email_master.append(emails[x].value.split(" ")[0])
                    #time_master = 0
                prev = trainers[x].value
            
            #for the last one since its basded on the change in value
            prev = retard_corrector(prev)
            if prev is None:
                continue
            day_master[prev.split(" ")[0]].append([hour_master, email_master])
            
    
    
    for x in list(day_master.keys())[1:]:
        email_address = ''     
        email_body = day_master['day']
        email_address = email_list[x]
        for y in list(day_master[x]):
            email_body += '\n' + y[0]
            for z in y[1]:
                email_body += '\n' + z
        if email_body is None:
            continue
        send_mail(email_address, email_body)
    
    master_list.append(day_master)
                
                    
                
        
        
        
        
        
        
        
        
        
        
        
        
        
