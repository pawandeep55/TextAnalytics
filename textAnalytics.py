# -*- coding: utf-8 -*-
import openpyxl



pCount=0
nCount=0
neuCount=0
positive=['love','good','nice','intelligent','smart','favourite']
negative=['bad','nincompoop']
neutral=[];
path="E:\
college\\bvicam\\4)sem\\dwdm\\text analytics\\cleansed(mail)\\shamsher_clean_data.xlsx"
wb=openpyxl.load_workbook(path)
print(wb.get_sheet_names())
sheet=wb.get_sheet_by_name('Sheet1')

input=sheet['B1'].value
splitStrings=input.split()
print splitStrings
for word in splitStrings:
    #print i
    if word in positive:
        pCount=pCount+1
    elif word in negative:
        nCount+=1
    elif word in neutral:
        neuCount=neuCount+1

if pCount>nCount:
    if pCount>neuCount:
        print 'tweet is postive'
    else:
        print 'tweet is neutral'
elif nCount>pCount:
    if nCount>neuCount:
        print 'tweet is negative'
    else:
        print 'tweet is neutral'
else:
   print 'not able to determine'
