# -*- coding: utf-8 -*-

input=raw_input('enter line for text analytics ')
splitStrings=input.split()
print splitStrings

pCount=0
nCount=0
neuCount=0
positive=['good','nice','intelligent','smart','favourite']
negative=['bad','nincompoop']
neutral=[];

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
#else:
 #   print 'not able to determine'
