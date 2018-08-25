# -*- coding: utf-8 -*-
"""
Created on Mon May 21 13:30:57 2018

@author: zachary.shaver
"""

from openpyxl import load_workbook

wb = load_workbook('//wtooleast.nosc.na.baesystems.com/Network_Documentation/Enterprise_Documentation/Voice/FedRAMP Migration 2018/Scripts/put_share_point_here/sharepoint.xlsx')

sheet = wb.get_sheet_by_name(wb.get_sheet_names()[0])

sch_method = sheet['L']
mig_status = sheet['K']

print(sch_method[1].value)
print(mig_status[1].value)

a = 0
b = 0
c = 0

for x in range(1, len(mig_status)):
    #print(mig_status[x].value, sch_method[x].value)
    if mig_status[x].value is None or sch_method[x].value is None:
        continue
    elif ("manual" in sch_method[x].value.lower() or 'ilearn' in sch_method[x].value.lower()) and "yes" in mig_status[x].value.lower():
        a+=1
    elif "fedramp" in sch_method[x].value.lower() and "yes" in mig_status[x].value.lower():
        b+=1
    elif ("ea" in sch_method[x].value.lower() or 'exec' in sch_method[x].value.lower()) and "yes" in mig_status[x].value.lower():
        c+=1

print("Interum fedRAMP", b, "Enterprise Users:", a, "Execs GG16", c)