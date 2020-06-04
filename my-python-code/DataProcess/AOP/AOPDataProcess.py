import openpyxl
import re

source = "source.xlsx"
wb = openpyxl.load_workbook(source)
ws = wb.worksheets[0]

print("Checking start...")

# 校验内容是否是从第四行开始

flag_AccMgr = str(ws.cell(3, 1).value) != "AccMgr"
flag_CustName = str(ws.cell(3, 2).value) != "Cust Name"
flag_CustCode = str(ws.cell(3, 3).value) != "Cust Code"
flag_FCG = str(ws.cell(3, 4).value) != "FCG"
flag_Plant = str(ws.cell(3, 5).value) != "Plant"
flag_Range = str(ws.cell(3, 6).value) != "Range"
flag_Model = str(ws.cell(3, 7).value) != "Model"
flag_SO = str(ws.cell(3, 8).value) != "SO"
flag_Product_ID = str(ws.cell(3, 9).value) != "Product ID"
flag_Config = str(ws.cell(3, 10).value) != "Config"
flag_Spec = str(ws.cell(3, 11).value) != "Spec"
flag_EngineFamily = str(ws.cell(3, 12).value) != "Engine Family"
flag_Application = str(ws.cell(3, 13).value) != "Application"
flag_Emission = str(ws.cell(3, 14).value) != "Emission"

if (flag_AccMgr and flag_CustName and flag_CustCode and flag_FCG \
        and flag_Plant and flag_Range and flag_Model and flag_SO
        and flag_Product_ID and flag_Config and flag_Spec \
        and flag_EngineFamily and flag_Application and flag_Emission):
    print("数据不是从第四行开始的")
else:
    if flag_AccMgr:
        print("检查三行A列是否为AccMgr,或列名是否包含空格")
    if flag_CustName:
        print("Cust Name")
    if flag_CustCode:
        print("Cust Code")
    if flag_FCG:
        print("FCG")
    if flag_Plant:
        print("Plant")
    if flag_Range:
        print("Range")
    if flag_Model:
        print("Model")
    if flag_SO:
        print("SO")
    if flag_Product_ID:
        print("Product ID")
    if flag_Config:
        print("Config")
    if flag_Spec:
        print("Spec")
    if flag_EngineFamily:
        print("Engine Family")
    if flag_Application:
        print("Application")
    if flag_Emission:
        print("Emission")

# 除了A-N列之外列数必须是13 整数倍
# 最大列号码
max_column = ws.max_column
max_row = ws.max_row
if (max_column - 14) % 13 != 0:
    print("请检查列数！！！")

# 检查内容是否为数字或者空
ws_value = ws.iter_rows(min_col=15, max_col=ws.max_column, min_row=4, max_row=max_row)
print(ws_value)

print([ (v.value for v  in row) for row in ws_value ])



# r = 1
# while max_row - 3:
#     print(ws_value)
#     r += 1

# s = 'w213'
# pattern = "^[-+]?(([0-9]+)([.]([0-9]+))?|([.]([0-9]+))?)$"
# print(re.match(pattern,s))
# if re.match(pattern,s) is None:
#     print("不通过")
