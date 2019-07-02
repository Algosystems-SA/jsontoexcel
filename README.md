# jsontoexcel
Converts a json file to csv and excel format.

It recursivelly explores the json file and flattens its data to output so it can be viewed as excel/csv

Takes a json file as input and converts it to csv and xlsx files. It recursivelly explores the json file's
It creates labels which correspond to the keys of the items parsed
It concatenates the full path tree json of each object to the label's name
structure and flattens the output in 2 dimensions so that the data can be converted to csv / xlsx format
you can use the -v flag for verbose output
Script call definition:
python jsontoexcel.py <Path to json> <Path and output files names(optional)>
  
# Usage
The first argument to specify will be the input json file the second is the base name used for the creation of the excel and csv file. If the second argument is not specified the json file's name will be used.
  
  This will output two files named: myfile.csv myfile.xlsx
  - python jsontoexcel.py myfile.json 
  
  Whereas this will output two files named: output.csv output.xlsx
  - python jsontoexcel.py myfile.json output - 
  
  you can place the -v argument anywhere eg:
  
  - python jsontoexcel.py myfile.json -v output
  - python jsontoexcel.py -v myfile.json output
  - python jsontoexcel.py myfile.json -v 
