# Sec25: excel to pdf

# To create a pattern for the file paths
import glob

# to read data which is .xlsx or .csv files
import pandas as pd

# to generate pdf
from fpdf import FPDF

# Creating the path files pattern , we specify directory and extension
# we will have the invoices files as a list
pathfiles = glob.glob("invoices/*.xlsx")
# to see the files
print(pathfiles)

# to read  the data of each file in a list
# and store it in variable
for filepath in pathfiles:
    # print the files inside the pathlist
    print(filepath)
    # we read data for each iteration filepath and get data for each excel file
    datafile = pd.read_excel(filepath, sheet_name="Sheet 1")
    # print the contents of each excel file
    print(datafile)

