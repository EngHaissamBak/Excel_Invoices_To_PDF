# sec 26 Excel to PDF
# continuing the sec 25 exercise

# To create a pattern for the file paths
import glob

# to read data which is .xlsx or .csv files
import pandas as pd

# to generate pdf
from fpdf import FPDF

# to extract a part of a file name we import pathlib library (Path function)
from pathlib import Path

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

    # to create a pdf object for each file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    # to add a page for this object
    pdf.add_page()

    # to write the Invoice nr. 10001 and date 2023.1.18

    # to get the part 10001  and 2023.1.18 from a file name  eg: fname = "10001-2023.1.18.xlsx"
    filename = Path(filepath).stem  # this removes the extension .xlsx of the filepath and takes only the stem name
    # --> filename = 10001-2023.1.18
    splitted_data = filename.split("-")  # splitting data on - and put it in data list
    # ['10001', '2023.1.18'}
    invoice_nr = splitted_data[0]    # take the element at index 0 = 10001
    date = splitted_data[1]          # take the element at index 1 = 2023.1.18
    # or another way to get invoice nr and date through one line (that  saves three lines of coding)
    # invoice_nr , date = filename.split("-")
    # this will split the filename that is without extension on - and put the 1st part in invoice_nr
    # and 2nd part in date

    # printing to check
    print(invoice_nr)
    print(date)

    pdf.set_font(family="Times",size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr} ", ln=1,border=0)

    # to write the date : Date 2023.1.18
    pdf.set_font(family="Times",size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date {date} ", ln=1,border=0)

    # generating pdf file for each excel file
    pdf.output(f"PDFS/{filename}.pdf")




