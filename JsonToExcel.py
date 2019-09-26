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
#   It creates labels which correspond to the keys of the items parsed
#   It concatenates the full path tree json of each object to the label's name
#   structure and flattens the output in 2 dimensions so that the data can be converted to csv / xlsx format
#   you can use the -v flag for verbose output
#	Script call definition:
#   python jsontoexcel.py <Path to json> <Path and output files names(optional)>
#
# Usage Sample:
#   python jsontoexcel.py myfile.json - will output two files: myfile.csv myfile.xlsx
#   Whereas : python jsontoexcel.py myfile.json output - will output two files: output.csv output.xlsx
#
#   https://opensource.org/licenses/MIT
#
# Copyright (c) George Panou panou.g@gmail.com https://github.com/george-panou
#

import sys
import json
import csv
import xlsxwriter

out = []
arg=""
i=0
verbose=False
#parse input parameters
if len(sys.argv) == 4:
    for i,arg in enumerate(sys.argv):
        if("-v" in arg):
            sys.argv.pop(i)
            verbose = True
    if(verbose):
        fileDir = str(sys.argv[1])
        outFile = str(sys.argv[2])
    else:
        print("Correct usage is : python jsontoexcel.py <Path to json> <Path and output files' base name(optional)>")
        print("Example : python jsontoexcel.py myfile.json")
        print("will output two files: myfile.csv myfile.xlsx\n")
        print("Whereas : python jsontoexcel.py myfile.json output")
        print("will output two files: output.csv output.xlsx")
        print("you can use the -v flag for verbose output")

elif len(sys.argv) == 3:
    for i,arg in enumerate(sys.argv):
        if("-v" in arg):
            sys.argv.pop(i)
            verbose = True
    if(verbose):
        fileDir = str(sys.argv[1])
        outFile = "./" + str(sys.argv[1]).split(".")[0]
    else:
        fileDir = str(sys.argv[1])
        outFile = str(sys.argv[2])

elif len(sys.argv) == 2:
    for i,arg in enumerate(sys.argv):
        if("-v" in arg):
            print("Correct usage is : python jsontoexcel.py <Path to json> <Path and output files' base name(optional)>")
            print("Example : python jsontoexcel.py myfile.json")
            print("will output two files: myfile.csv myfile.xlsx\n")
            print("Whereas : python jsontoexcel.py myfile.json output")
            print("will output two files: output.csv output.xlsx")
            print("you can use the -v flag for verbose output")
            sys.exit(-1)
    fileDir = str(sys.argv[1])
    outFile = "./"+str(sys.argv[1]).split(".")[0]

else:
    print ("Correct usage is : python jsontoexcel.py <Path to json> <Path and output files' base name(optional)>")
    print("Example : python jsontoexcel.py myfile.json")
    print("will output two files: myfile.csv myfile.xlsx\n")
    print("Whereas : python jsontoexcel.py myfile.json output")
    print("will output two files: output.csv output.xlsx")
    print("you can use the -v flag for verbose output")

    sys.exit(-1)

#flattens a tree object consisted of dictionaries and lists
def flatten_json(y):
    print("flattening json file recursivelly")
    list2 = []
    labels = []
    depth = []
    global count
    count = 0

#flatten each row of the root list
    if type(y) is dict:
        for j in y.values() :
            #print(j)
            out,lbl,cnt=flatten(j,' ')
            if verbose:
                print("Sub tree:" + str(out))
            depth.append(cnt)
            labels.append(lbl)
            #print(out)
            list2.append(out)

    elif isinstance(y, list):
        for j in y :
            #print(j)
            out,lbl,cnt=flatten(j,' ')
            if verbose:
                print("Sub tree:" + str(out))
            depth.append(cnt)
            labels.append(lbl)
            #print(out)
            list2.append(out)
    label=[]

    #find max path in json tree
    label.append( max(labels, key=len))
    if verbose:
        print("labels:"+str(label))
    list2 = label + list2
    if verbose:
        print (list2)
    return (list2)

labels = []

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
print("Loading json file")
with open(fileDir,encoding = 'utf-8', newline='') as file:
    data = file.read().replace('\n', '')
all_data = json.loads(data)
print(all_data)
global count

#flatten data
flat = flatten_json(all_data)

#create csv with flattened data
print("Saving data as "+outFile+".csv")
data_csv = open(outFile+".csv", 'w',newline='')
csvwriter = csv.writer(data_csv)
data_csv.write('SEP=,\n')
for row in flat :
    csvwriter.writerow(row)


#save data as xlsx
print("Saving data as "+outFile +'.xlsx')
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
