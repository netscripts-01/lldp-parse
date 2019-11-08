#!/usr/bin/env python

import re
import os
import xlwt
from glob import glob

outfilename = "Building_1.xls"
stringToMatch = "SysName"

wb = xlwt.Workbook()                          #create an Excel workbook

for name in glob('*/'):                       #for each directory folder name
    path = name                               #use folder name retrieved as filepath
    sheetname = name[:-1]                     #trim the name    
    ws = wb.add_sheet(sheetname) 
    ws.write(0,0, 'DOSE')
    ws.write(0,1, 'SYSNAME')
    ws.write(0,2, 'PORTDESCR') 
    wb.save(outfilename)
    n = 0                                     #reset incrementing row value for spreadsheet to zero

    for filename in os.listdir(path):                                    #for each text file in the specified path
        with open(os.path.join(path, filename)) as filetoread:           #concatenate dir name and file name into a path
            if filename.endswith(".txt"):                                #check file is txt file
                fd = open(os.path.join(path, filename))                  #create fd as a variable for the open file
                with open(os.path.join(path, filename)) as openfile:     
                     n = n+1                                             #set variable to iterate up sheet rows
                     if (stringToMatch) in openfile.read():              #if given string is present, get info
                         for line in fd:
                             match = re.search(r'SysName:      (.*)', line)
                             if match:
                                 sysname = match.group(1)
                                 print(sysname)
                    
                             match = re.search(r'PortDescr:    (.*)', line)
                             if match:
                                 portdescr = match.group(1)
                                 print(portdescr)

                     else:                                               #if given string is not present return 0000
                          sysname = ("0000")
                          print(sysname)
                          portdescr = ("0000")
                          print(portdescr)

                     dosename = filename[:-4]                            #get socket name from txt file name

                     ws.write((n),0, ''.join(dosename))                  #write values, iterating down rows
                     ws.write((n),1, ''.join(sysname))
                     ws.write((n),2, ''.join(portdescr))
                     wb.save(outfilename)


