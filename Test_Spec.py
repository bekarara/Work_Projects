import xlsxwriter
import pandas as pd
from docx import Document
from openpyxl import Workbook
from openpyxl import load_workbook
#===========Inputs=======================================#
word_doc = "D:\\Downloads\\TA8X.0100550.S06.00EN,01.00 QNGR AoE Lab Test Specification.docx"
output = "C:\\Users\\xrbek\\Documents\\repo\\Output_Files\\test_results.xlsx"
# uses "NGR_TC_SI_LAB" to filter for the targted tables.

#========================================================#
def convert_doc_table_to_excel(word_doc,output):
    # read the test spec document
    doc = Document(word_doc)

    #convert to df
    tables = []
    for table in doc.tables:
        num_rows = len(table.rows)
        num_cols = len(table.columns)
        df = pd.DataFrame(index=range(num_rows), columns = range(num_cols))

        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                df.iloc[i,j] = cell.text

        #keep cells with unique values only:
        df = df.apply(lambda x: x.drop_duplicates(keep = 'first'),axis=1)

        #filter for test case tables only:
        if "NGR_TC_SI_LAB" in (table.rows[0].cells[0].text or table.rows[1].cells[0].text):
            df.iloc[12,4] = 'Test Result'
            df.iloc[12,5] = 'Remarks'
            for i in range(13,df.shape[0]): #df.shape[0] gives number of rows
                df.iloc[i,0] = i-12
            tables.append(pd.DataFrame(df))

    #write to excel
    writer = pd.ExcelWriter(output, engine = 'xlsxwriter')
    for df in tables:
        # write dataframes to excel
        df.to_excel(writer, sheet_name = df.iloc[0][0][18:22])
        
        #===========formatting===================================#
        workbook = writer.book
        #name sheets with test case ID
        worksheet = writer.sheets[df.loc[0][0][18:22]]
        #wrap text and set column width
        format_text_wrap = workbook.add_format({'text_wrap' : True})
        writer.sheets[df.loc[0][0][18:22]].set_column(1,6,40,format_text_wrap)
        # colour header row
        format_cell_colour = workbook.add_format({'bold' : True, 'bg_color' : 'silver'})
        #freeze pane
        worksheet.freeze_panes(14,0)
        #group the rows that are not test steps
        for freeze_row in range(2,13):
            worksheet.set_row(freeze_row, None, None, {'level' : 2, "hidden": True})
        worksheet.set_row(freeze_row+2, None, None,{'collapsed' : True})
        #colour header cells
        worksheet.set_row(13,None,format_cell_colour)
        
    writer.close()
    return

convert_doc_table_to_excel(word_doc,output)