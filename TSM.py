#Writer: Michael Guarino
#TSM Report
#DO NOT TOUCH :0)

import easygui
import extract_msg
import re
import io
import openpyxl
import datetime
import os

#File open and message/date extraction
f = easygui.fileopenbox()
date = os.path.basename(f)[:-4]
msg = extract_msg.Message(f)
msg_report = msg.body
#Locates nodes
s = msg_report.find("Online Nodes")
e = msg_report.find("Nodes offline")

nodes = msg_report[s:e]
iterString = io.StringIO(nodes)

#Pulls nodes and statuses
status = []
node = []
for line in iterString:
    if line[0] == 'F' or line[0] == 'M':
        temp = re.split(r'\t+', line)
        status.append(str(temp[0])[0])
        node.append(temp[1])

#Opens Excel sheet for editing
from openpyxl import load_workbook
wb = load_workbook('Z:\IT\Ops Shared\Daily Checklist\TSM Report Log.xlsx')
ws = wb['2019']

#Finding date column in sheet
colCurr = 0
dateFound = False
for col in ws.iter_cols(min_row=1, max_col=370, max_row=1):
    for cell in col:
        if cell.value != "Node Name":
            temp = str(cell.value)[:-9]
            year, month, day = temp.split('-')
            temp = year + '-' + month + '-' + day
            if str(temp) == str(date):
                colCurr = cell.column
                print("Column (" + date + ") found")
                dateFound = True

if dateFound == False:
    print("Error: Date (" + date + ") not found.")
        
#Finds node and inputs status
rowCurr = 0
cnt = 0
nodesNotFound = []
notFoundStatus = []
if colCurr != 0:
    for n in node:
        nodeFound = False
        for row in ws.iter_rows(min_col=1, max_row=750, max_col=1):
            for cell in row:
                if str(n) == str(cell.value).strip():
                    nodeFound = True
                    rowCurr = cell.row        
                    target = ws.cell(row=rowCurr, column=colCurr)
                    prev = ws.cell(row=rowCurr, column=colCurr-1)
                    if prev.value is not None:
                        prevState = str(prev.value)
                        failState, failCount = prevState.split('-')
                        if failState == status[cnt]:
                            newCount = int(failCount) + 1
                            target.value = str(str(status[cnt]) + '-' + str(newCount))
                        else:
                            target.value = status[cnt] + '-1'
                    else:
                        target.value = status[cnt] + '-1'  
                    
        if nodeFound == False:
            nodesNotFound.append(n)
            notFoundStatus.append(status[cnt])
            
        cnt = cnt+1
        
print("Not found in sheet: " + str(nodesNotFound))
print("Inserting into sheet...")

ins = -1
notFoundCnt = 0
for n in nodesNotFound:
    posFound = False
    for row in ws.iter_rows(min_col=1, max_row=750, max_col=1, min_row=2):
        for cell in row:
            if str(n) < str(cell.value) and posFound == False:
                ins = cell.row
                posFound = True
    ws.insert_rows(ins)
    ws.cell(row=ins, column=1).value = str(n)
    ws.cell(row=ins, column=colCurr).value = str(notFoundStatus[notFoundCnt] + "-1")
    notFoundCnt = notFoundCnt+1


print("This what God feel like (done)")  
wb.save('Z:\IT\Ops Shared\Daily Checklist\TSM Report Log.xlsx')
    
#print(status)
#print(node)
