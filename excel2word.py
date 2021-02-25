import xlrd
from docx import Document
from docx.oxml.ns import qn
import shutil
import os
import re


for dataset in ["unlabel_data_存款准备金类10000","unlabel_data_工业10000","unlabel_data_价格指数10000"]:
    print(dataset)
    data = xlrd.open_workbook("./data_excel/"+dataset+".xlsx")
    # unlabel_data_工业10000
    # unlabel_data_价格指数10000
    table = data.sheets()[0]
    # names = data.sheet_names() 
    # print(names)
    nrows = table.nrows 
    # range(1,nrows)
    for i in range(1,nrows):
        row_detail = table.row_values(i)
        # print(row_detail)

        row_detail[2] = re.sub(r'\s+', "\n\t", row_detail[2])
        row_detail[1] = re.sub(r'\/', ' ', str(row_detail[1]))

        document = Document()
        document.add_paragraph(row_detail[2])
        document.styles['Normal'].font.name = u'宋体'
        document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
        try:
            document.save(row_detail[1]+".docx")
        except OSError:
            print(dataset+" " + str(i) + " name too long")
            row_detail[1] = row_detail[1][10]
            document.save(row_detail[1]+".docx")

        try:
            shutil.move(row_detail[1]+".docx",dataset)
        except shutil.Error:
            os.remove(row_detail[1]+".docx")
            print(dataset+" " + str(i)+" exist")


        # print(row_detail)




