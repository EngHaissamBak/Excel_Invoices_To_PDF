# To create a pattern for the file paths
import glob

# to read data which is .xlsx or .csv files
import pandas as pd

# to generate pdf
from fpdf import FPDF

# to extract a part of a file name we import pathlib library (Path function)
from pathlib import Path

# Creating the path files pattern , we specify directory and extension
pathfiles = glob.glob("invoices/*.xlsx")
# to see the files
print(pathfiles)

# to read  the data of each file in a list
for filepath in pathfiles:
    print(filepath)
    # to create a pdf object for each file and add a page
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # to write the Invoice nr. 10001 and date 2023.1.18
    filename = Path(filepath).stem  # this removes the extension .xlsx of the filepath and takes only the stem name
    # --> filename = 10001-2023.1.18
    splitted_data = filename.split("-")  # splitting data on - and put it in data list
    # ['10001', '2023.1.18'}
    invoice_nr = splitted_data[0]    # take the element at index 0 = 10001
    date = splitted_data[1]          # take the element at index 1 = 2023.1.18

    # printing to check
    print(invoice_nr)
    print(date)

    pdf.set_font(family="Times",size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr} ", ln=2,border=0)

    # to write the date : Date 2023.1.18
    pdf.set_font(family="Times",size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date {date} ", ln=2, border=0)

    # adding cells to the pdf and the data from excel files
    datafile = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(datafile)

    # Add Header
    col = list(datafile.columns)
    col = [item.replace("_", " ").title() for item in col]
    print(col)
    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt=col[0], align="C", border=1)
    pdf.cell(w=50, h=8, txt=col[1], align="C", border=1)
    pdf.cell(w=50, h=8, txt=col[2], align="C", border=1)
    pdf.cell(w=30, h=8, txt=col[3], align="C", border=1)
    pdf.cell(w=30, h=8, txt=col[4], align="C", border=1, ln=1)

    # Add Rows generating the table cells rows in pdf
    for index, row in datafile.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), align="C", border=1)
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), align="C",border=1)
        pdf.cell(w=50, h=8, txt=str(row["amount_purchased"]),align="C", border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), align="C",border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), align="C", ln=1, border=1)

    # Calculate total price and add it to a new row in last cell under total price
    total_sum = datafile["total_price"].sum()
    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=160, h=8, txt="TOTAL PRICE IN EUROS IS: ", align="L", border=1)  # cell 4 (empty)
    pdf.cell(w=30, h=8, txt=str(total_sum), align="C", ln=1, border=1)  # cell 5

    # To write “ The total due amount is 339 Euros “after the table
    pdf.set_font(family="Times", style="B", size=20)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=8, txt=f"The total due amount is {total_sum} Euros.", align="L", ln=1, border=0)

    # Add company name and logo
    pdf.set_font(family="Times", style="B", size=20)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=70, h=10, txt="BIMATIC EUROPE", align="L", border=0)
    pdf.image("logooo.jpg", w=10, h=10)

    # generating pdf file for each excel file
    pdf.output(f"PDFS/{filename}.pdf")




