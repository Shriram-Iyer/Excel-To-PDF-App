import pandas as pd
from fpdf import FPDF
from pathlib import Path


def exceltopdf(filepaths, dest_dir):
    for filepath in filepaths:
        print(filepath)
        if ".xls" in filepath or ".xlsx" in filepath:
            df = pd.read_excel(filepath, sheet_name='Sheet1')
            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_font(family='Times', size=16, style='B')
            filename = Path(filepath).stem
            invno, date = filename.split('-')
            pdf.cell(w=50, h=8, txt=f"Invoice No: {invno}", ln=1)
            pdf.cell(w=50, h=8, txt=f"Invoice Date: {date}", ln=1)
            columns = [item.replace('-', ' ').title() for item in df.columns]
            pdf.set_font(family='Times', size=10, style='B')
            pdf.cell(w=30, h=8, txt=str(columns[0]), border=1)
            pdf.cell(w=70, h=8, txt=str(columns[1]), border=1)
            pdf.cell(w=30, h=8, txt=str(columns[2]), border=1)
            pdf.cell(w=30, h=8, txt=str(columns[3]), border=1)
            pdf.cell(w=30, h=8, txt=str(columns[4]), border=1, ln=1)
            for index, row in df.iterrows():
                pdf.set_font(family='Times', size=10)
                pdf.cell(w=30, h=8, txt=str(row['Product_id']), border=1)
                pdf.cell(w=70, h=8, txt=str(row['Product_name']), border=1)
                pdf.cell(w=30, h=8, txt=str(row['quantity']), border=1)
                pdf.cell(w=30, h=8, txt=str(row['unit_price']), border=1)
                pdf.cell(w=30, h=8, txt=str(row['Amount']), border=1, ln=1)
            pdf.set_font(family='Times', size=10)
            pdf.cell(w=30, h=8, txt=str(''), border=1)
            pdf.cell(w=70, h=8, txt=str(''), border=1)
            pdf.cell(w=30, h=8, txt=str(''), border=1)
            pdf.cell(w=30, h=8, txt=str(''), border=1)
            pdf.cell(w=30, h=8, txt=str(df['Amount'].sum()), border=1, ln=1)

            pdf.output(dest_dir + "/" + filename + ".pdf")
            return True
        else:
            return False
