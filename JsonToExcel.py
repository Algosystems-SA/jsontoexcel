# Takes a json file as input and converts it to csv and xlsx files
#
# Prerequisites:
#   - Python Libraries
#     - xlsxwriter
#
#   You can install the libraries using pip:
#    - pip install XlsxWriter
#
# Version 0.1
#   Takes a json file as input and converts it to csv and xlsx files. It recursivelly explores the json file's
#   structure and flattens the output in 2 dimensions so that the data can be converted to csv / xlsx format
#	Script call definition:
#   python jsontoexcel.py <Path to json> <Path and output files names(optional)>
#   http://www.opensource.org/licenses/gpl-2.0.php
#
# Copyright (c) George Panou panou.g@gmail.com https://github.com/george-panou
#

import sys
import json
import csv
import xlsxwriter


out = []

if len(sys.argv) == 3:
    fileDir = str(sys.argv[1])
    outFile = str(sys.argv[2])
elif len(sys.argv) == 2:
    fileDir = str(sys.argv[1])
    outFile = "./converted"
else:
    print ("Correct usage is : python jsontoexcel.py <Path to json> <Path and output files names(optional)>")
    sys.exit()

#flattens a tree object consisted of dictionaries and lists
def flatten_json(y):
    list2 = []
    labels = []
    depth = []
    global count
    count = 0

#flatten each row of data
    for j in y :
        #print(j)
        out,lbl,cnt=flatten(j,' ')
        depth.append(cnt)
        labels.append(lbl)
        #print(out)
        list2.append(out)
    label=[]

    #find max path in json tree
    label.append( max(labels, key=len))
    #print(label)
    list2 = label + list2
    #print (list2)
    return (list2)

labels = []
global depth
depth=[]

#explore a tree with recursion and flatten to list
def flatten(x,name):
    out=[]
    label=[]
    count=0
    if type(x) is dict:
        for a in x:
            tmp,nm,cnt=flatten(x[a], name + a + '/')
            out+=tmp
            label+=nm
            count+=cnt
    elif isinstance(x, list):
        i = 0
        for a in x:
            tmp,nm,cnt=flatten(a, name + str(i) + '/')
            out+=tmp
            label += nm
            count+=cnt
            i += 1
    else:
        count += 1
        out.append(x)
        label.append(name)

    return out,label,count


#open json file
with open(fileDir,encoding = 'utf-8', newline='') as file:
    data = file.read().replace('\n', '')

all_data = json.loads(data)
global count

#create csv with flattened data
data_csv = open(outFile+".csv", 'w',newline='')
csvwriter = csv.writer(data_csv)
data_csv.write('SEP=,\n')

flat = flatten_json(all_data['data'])

for row in flat :
    csvwriter.writerow(row)


#save data as xlsx
workbook = xlsxwriter.Workbook(outFile +'.xlsx',)
worksheet = workbook.add_worksheet()

bold = workbook.add_format({'bold': True})

for r, row in enumerate(flat):
    for c, col in enumerate(row):
        if r==0:
            worksheet.write(r, c, col, bold)
        else:
            worksheet.write_string(r, c, str(col))

workbook.close()

print("Successfully created files:"+outFile +'.xlsx , ' + outFile+".csv" )
