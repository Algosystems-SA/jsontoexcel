import json
import sys
import xlsxwriter
import csv

out = []

if len(sys.argv) == 3:
    fileDir = str(sys.argv[1])
    outFile = str(sys.argv[2])
elif len(sys.argv) == 2:
    fileDir = str(sys.argv[1])
    outFile = "./converted"
else:
    print ("Correct usage is : <./jsontoexcel.py> <Path to json> <Path for output files (optional)>")
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
