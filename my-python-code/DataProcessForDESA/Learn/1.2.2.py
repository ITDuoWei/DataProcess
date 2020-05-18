import openpyxl
wb = openpyxl.load_workbook("我的工作簿.xlsx",data_only=True)
print(wb)
# 另存为
wb.save("我的工作簿副本.xlsx")

# str = "wei"
# str1 = "weidfd"
# print(str in str1)
#
# l = ['a','b','c','d']
# l1 = [x for x in l if x in ['a','b']]
# print(l1)