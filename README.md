# jsontoexcel
Converts a json file to csv and excel format.

It recursivelly explores the json file and flattens its data to output so it can be viewed as excel/csv

Takes a json file as input and converts it to csv and xlsx files. It recursivelly explores the json file's data (nodes)
It creates labels which correspond to the keys of the items parsed
It concatenates the full path tree json of each object to the label's name structure  and flattens the output in 2 dimensions so that the data can be converted to csv / xlsx format
You can use the -v flag for verbose output


  
# Usage
Script call definition:
python jsontoexcel.py <Path to json> <Path and output files names(optional)>
  
The first argument you specify will be the input json file the second is the base name used for the creation of the excel and csv file. If the second argument is not specified the json file's name will be used.
  
  This will output two files named: myfile.csv myfile.xlsx
  - python jsontoexcel.py myfile.json 
  
  Whereas this will output two files named: output.csv output.xlsx
  - python jsontoexcel.py myfile.json output  
  
  you can place the -v argument anywhere eg:
  
  - python jsontoexcel.py myfile.json -v output
  - python jsontoexcel.py -v myfile.json output
  - python jsontoexcel.py myfile.json -v 

# Known Issues

There are no issues so far, please communicate with me if you run into any, so that I can improve the script.

# Getting Help

If you face any issues please create an issue in the github built-in issue tracker

# Contributing

Please contact me if you want to contribute to this project and I will give you further instructions. 
