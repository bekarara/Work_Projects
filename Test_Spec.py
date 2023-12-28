import xlsxwriter
import pandas as pd
from docx import Document
from openpyxl import Workbook
from openpyxl import load_workbook
#======================================================Start of script
path = "D:\\Downloads\\TA8X.0100550.S06.00EN,01.00 QNGR AoE Lab Test Specification.docx"
output = 'test_results.xlsx'
# read the test spec document
doc = Document(path)

#convert to df
tables = []
for table in doc.tables: #table is docx.table.Table type
    num_rows = len(table.rows)
    num_cols = len(table.columns)
    df = pd.DataFrame(index=range(num_rows), columns = range(num_cols))

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            df.iloc[i,j] = cell.text
    df = df.apply(lambda x: x.drop_duplicates(keep = 'first'),axis=1)
    if "NGR_TC_SI_LAB" in (table.rows[0].cells[0].text or table.rows[1].cells[0].text):
        tables.append(pd.DataFrame(df))

#write to excel
writer = pd.ExcelWriter(output, engine = 'xlsxwriter')
for df in tables:
    # write dataframes to excel
    df.to_excel(writer, sheet_name = df.loc[0][0][18:22])
    
    #===========formatting===================================#
    workbook = writer.book
    #name sheets with test case ID
    worksheet = writer.sheets[df.loc[0][0][18:22]]
    #wrap text and set column width
    text_format = workbook.add_format({'text_wrap' : True})
    writer.sheets[df.loc[0][0][18:22]].set_column(1,6,40,text_format)
    #freeze pane
    worksheet.freeze_panes(12,0)
    #group the rows that are not test steps
    for freeze_row in range(2,13):
        worksheet.set_row(freeze_row,None, None, {'level' : 2, "hidden": True})
    worksheet.set_row(freeze_row+2,None, None, {'collapsed' : True})
writer.close()
