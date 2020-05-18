import openpyxl
wb = openpyxl.load_workbook("我的工作簿.xlsx")
# 获取活动工作表
ws = wb.active
# 以索引值得方式获取工作表
ws1 = wb.worksheets[0]
# 以名称的方式
ws2 = wb['Sheet']
print(ws2)

# for sh in wb.worksheets:
#     print(sh)

print(wb.sheetnames)

wb.worksheets[0].title='demo'
wb.save("我的工作簿.xlsx")