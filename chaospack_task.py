import os               #deal with the operating system
import pandas as pd     #work with data sets
import csv              #import and export format for spreadsheet and databases
import PyPDF2           #pure-python PDF library
from PyPDF2 import PdfMerger

path = 'c:/Users/Jinius/Downloads/Employee Registration 2020/NewEmployees2020Clean'

#check whether directory already exists
if not os.path.exists(path):                            #if the directory does not exist
    #create new single directory
    os.mkdir(path)
    print(f"Folder {path} successfully created!")
else:                                                   #if the directory exists
    print(f"Folder {path} already exists.")

D = r"c:/Users/Jinius/Downloads/Employee Registration 2020/NewEmployees2020Clean"     #destination directory

#create a result.csv file for PART C
result_file = "c:/Users/Jinius/Downloads/Employee Registration 2020/NewEmployees2020Clean/result.csv"

f = open(result_file, "w")      #will create a file if the specified file does not exist

# pdf_files = []

#PART A
print("Before copying file:", os.listdir(D))    #to display what is inside the directory before copying files

month = ['Jan', 'Feb', 'March', 'Apr',           #month variables
         'May', 'Jun', 'Jul', 'Aug',
         'Sept', 'Oct', 'Nov', 'Dec']

for i in range(len(month)):                     #loop 12 times from month[0] to month[11]
    
    #path contains the path of the source file
    #format() formats the specifies value and insert them inside the string's placeholder
    src = "c:/Users/Jinius/Downloads/Employee Registration 2020/Employee Registration 2020/Employee Registration {} 2020/newjoiners.csv".format(month[i])  
    
    if not os.path.exists(src):     #if the path does not have newjoiners.csv file
        print(f"{month[i]} does not have any new employees.")
        with open(result_file, 'a+') as t:
            t.write("{},".format(month[i]))      #write 0 for the non-existing files (for the month that has no new employees)
            t.write('0')
            t.write("\n")

    else:
        #dest contains the path of the destination file
        dest = "c:/Users/Jinius/Downloads/Employee Registration 2020/NewEmployees2020Clean/newjoiners_{}2020.csv".format(month[i])

        header = ['First name', 'Last name', 'Email', 'Employee ID']

        #PART B
       
        with open(src, 'r') as f:
            with open(dest, 'w', newline='') as t:            #add newline to avoid adding empty line after the header
                dw = csv.DictWriter(t, fieldnames=header)     #write the header before writing the file (only suitable for new csv file)
                dw.writeheader()
                
                for lines in f:
                    new_lines = lines.replace(";", ",")       #change the delimiter from a semicolon into a comma
                    t.write(new_lines)
          
        #PART C
        #read the csv files
        # reader = csv.reader(open(dest))
        
        with open(dest, 'r') as f:
            heading = next(f)                          #start counting the no of rows after the header
            reader_obj = csv.reader(f)                 #read the csv files
            row = len(list(reader_obj))                #get the no of rows 
            with open(result_file, 'a+') as t:      
                #data is put inside a list 
                #the no of employees = no of rows                    
                t.write("{},".format(month[i]))     #write month in the first column
                t.write(str(row))                   #convert int to str
                t.write("\n")                       #put a spacing

    #BONUS TASK B
    #new PdfMerger instance is created for each month, preventing the previous months from being appended to subsequent months
    merger = PdfMerger()   
    pdf_files = []
    pdf_src = "c:/Users/Jinius/Downloads/Employee Registration 2020/Employee Registration 2020/Employee Registration {} 2020".format(month[i])
    os.chdir(pdf_src)    #change to the src directory

    pdf_dest = "c:/Users/Jinius/Downloads/Employee Registration 2020/NewEmployees2020Clean/AllNewEmployees_{}2020.pdf".format(month[i])

    if len(os.listdir(pdf_src)) == 0:
        continue
        # print(f"{month[i]} does not have new employees.")

    else:
        
        for filename in os.listdir(pdf_src):        #inside a directory of src
            if filename.endswith(".pdf"):           #if the filenames are ended with .pdf
                pdf_files.append(filename)          #append all the files together

        for pdf in pdf_files:                       #loop the appended files
            merger.append(pdf)                      #merge the files together
        
        # del pdf_files[:]                            #clear the list of the pdf files to avoid duplicating pdf in the subsequent months

        merger.write(pdf_dest)                      #write all data that has been merged to the given output file
        merger.close()                              #shut all file descriptors (input and output) and clear all memory usage


#write the Total at the last row of the data
column_names = ["Month", "No. of Employees"]        #forced to add headers so that Jan data can be retrieved
df = pd.read_csv(result_file, names=column_names)   #read the result.csv together with the headers
no_of_employees = df.iloc[:,1]                      #retrieve the column of no. of employees of each month
total = sum(no_of_employees)                        #add up to get the total number of employees
f = open(result_file, "a")                          #open the file (the file is closed at  with statement) and append the data to the existing result.csv
f.write("Total,")                                   #write "Total"
f.write(str(total))                                 #write down the number
f.close()                                           #close the file

print("After copying file:", os.listdir(D))         #to display what is inside the directory after copying files





#import shutil
#create duplicate of the file at the destination with the filename mentioned
#at the end of the destination path
#if a file with the same name does not exist in the destination
#a new file with the name mentioned is created
# dest = shutil.copyfile(src, dest)