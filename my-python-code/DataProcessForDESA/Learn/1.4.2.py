import openpyxl
wb = openpyxl.load_workbook("demo1.xlsx")
wb.copy_worksheet(wb['工资表']).title="工资表1月"
wb.remove(wb['Sheet'])
wb.save("demo01.xlsx")

