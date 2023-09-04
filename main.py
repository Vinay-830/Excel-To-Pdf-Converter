import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoice/*.xlsx")

for filepath in filepaths:
    # extract  invoice nr and date of invoice
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Generate PDF invoice
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr {invoice_nr}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date {date}", ln=1)
    pdf.cell(w=0, h=10, txt=" ", ln=1)

    # read excel file into data frame
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # display the title of invoice table
    column = list(df.columns)
    column = [ col.title().replace("_", " ") for col in column]
    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=column[0], border=1)
    pdf.cell(w=70, h=8, txt=column[1], border=1)
    pdf.cell(w=30, h=8, txt=column[2], border=1)
    pdf.cell(w=30, h=8, txt=column[3], border=1)
    pdf.cell(w=30, h=8, txt=column[4], border=1, ln=1)

    # display row data
    price = 0
    for index, row in df.iterrows():
        price += row["total_price"]

        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=10, txt=" ", ln=1)
    pdf.cell(w=50, h=8, txt=f"The total due amount is {price} Euros.", ln=1)
    pdf.cell(w=50, h=8, txt="Python How", ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
