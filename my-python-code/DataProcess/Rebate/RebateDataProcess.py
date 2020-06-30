import openpyxl
import datetime

table_desa = "destination.xlsx"
table_rebate = "rebate.xlsx"

# 获取工作簿
wb_desa = openpyxl.load_workbook(table_desa)
wb_rebate = openpyxl.load_workbook(table_rebate)

#
# 获取工作表
ws_desa = wb_desa.worksheets[0]
ws_rebate = wb_rebate.worksheets[0]

list_desa = (list(ws_desa.values))
list_MTBECE = []
list_TBECE = []

# 获取列表 [Team,LOB,EMISSION,Costume,Engine Family]
r = 2
while r <= ws_desa.max_row:
    list_MTBECE.append(
        [ws_desa.cell(r, 1).value, ws_desa.cell(r, 46).value, ws_desa.cell(r, 47).value, ws_desa.cell(r, 48).value,
         ws_desa.cell(r, 49).value,
         ws_desa.cell(r, 50).value])
    list_TBECE.append(
        [ws_desa.cell(r, 46).value, ws_desa.cell(r, 47).value, ws_desa.cell(r, 48).value, ws_desa.cell(r, 49).value,
         ws_desa.cell(r, 50).value])
    r += 1;

# print(list_MTBECE)
# print(list_TBECE)

listJan = ["Jan"]
listFeb = ["Feb"]
listMar = ["Mar"]
listApr = ["Apr"]
listMay = ["May"]
listJun = ["Jun"]
listJul = ["Jul"]
listAug = ["Aug"]
listSep = ["Sep"]
listOct = ["Oct"]
listNov = ["Nov"]
listDec = ["Dec"]
# 获取当前的月份
month = datetime.datetime.now().month
# 每个月的折扣数量
debatecount = 0

# 判断数据部分是否为空
rebate_data = ws_rebate['F':'Q']
list_rebate_data = [[r.value for r in rn][1:] for rn in rebate_data]
if list_rebate_data == [[None, None], [None, None], [None, None], [None, None], [None, None], [None, None],
                        [None, None], [None, None], [None, None], [None, None], [None, None], [None, None]]:
    Noneflag = True

# 循环处理rebate数据  Customer,Engine Family,Emission,LOB,Team
list_TBECE_rebate = []
j = 2
while j <= ws_rebate.max_row:
    list_TBECE_rebate = [ws_rebate.cell(j, 5).value, ws_rebate.cell(j, 4).value, ws_rebate.cell(j, 3).value,
                         ws_rebate.cell(j, 1).value,
                         ws_rebate.cell(j, 2).value]
    # 如果数据值全部为空,则说明首次使用，直接更新所有数据
    if Noneflag:
        debatecount = list_MTBECE.count(listJan + list_TBECE_rebate)
        ws_rebate.cell(j, 6, debatecount)

        debatecount = list_MTBECE.count(listFeb + list_TBECE_rebate)
        ws_rebate.cell(j, 7, debatecount)

        debatecount = list_MTBECE.count(listMar + list_TBECE_rebate)
        ws_rebate.cell(j, 8, debatecount)

        debatecount = list_MTBECE.count(listApr + list_TBECE_rebate)
        ws_rebate.cell(j, 9, debatecount)

        debatecount = list_MTBECE.count(listMay + list_TBECE_rebate)
        ws_rebate.cell(j, 10, debatecount)

        debatecount = list_MTBECE.count(listJun + list_TBECE_rebate)
        ws_rebate.cell(j, 11, debatecount)

        debatecount = list_MTBECE.count(listJul + list_TBECE_rebate)
        ws_rebate.cell(j, 12, debatecount)

        debatecount = list_MTBECE.count(listAug + list_TBECE_rebate)
        ws_rebate.cell(j, 13, debatecount)

        debatecount = list_MTBECE.count(listSep + list_TBECE_rebate)
        ws_rebate.cell(j, 14, debatecount)

        debatecount = list_MTBECE.count(listOct + list_TBECE_rebate)
        ws_rebate.cell(j, 15, debatecount)

        debatecount = list_MTBECE.count(listNov + list_TBECE_rebate)
        ws_rebate.cell(j, 16, debatecount)

        debatecount = list_MTBECE.count(listDec + list_TBECE_rebate)
        ws_rebate.cell(j, 17, debatecount)
    else:
        # 如果统计到的记录大于0则更新
        if list_TBECE.count(list_TBECE_rebate) > 0:
            if month == 1:
                debatecount = list_MTBECE.count(listJan + list_TBECE_rebate)
                ws_rebate.cell(j, 6, debatecount)
            if month == 2:
                debatecount = list_MTBECE.count(listFeb + list_TBECE_rebate)
                ws_rebate.cell(j, 7, debatecount)
            if month == 3:
                debatecount = list_MTBECE.count(listMar + list_TBECE_rebate)
                ws_rebate.cell(j, 8, debatecount)
            if month == 4:
                debatecount = list_MTBECE.count(listApr + list_TBECE_rebate)
                ws_rebate.cell(j, 9, debatecount)
            if month == 5:
                debatecount = list_MTBECE.count(listMay + list_TBECE_rebate)
                ws_rebate.cell(j, 10, debatecount)
            if month == 6:
                debatecount = list_MTBECE.count(listJun + list_TBECE_rebate)
                ws_rebate.cell(j, 11, debatecount)
            if month == 7:
                debatecount = list_MTBECE.count(listJul + list_TBECE_rebate)
                ws_rebate.cell(j, 12, debatecount)
            if month == 8:
                debatecount = list_MTBECE.count(listAug + list_TBECE_rebate)
                ws_rebate.cell(j, 13, debatecount)
            if month == 9:
                debatecount = list_MTBECE.count(listSep + list_TBECE_rebate)
                ws_rebate.cell(j, 14, debatecount)
            if month == 10:
                debatecount = list_MTBECE.count(listOct + list_TBECE_rebate)
                ws_rebate.cell(j, 15, debatecount)
            if month == 11:
                debatecount = list_MTBECE.count(listNov + list_TBECE_rebate)
                ws_rebate.cell(j, 16, debatecount)
            if month == 12:
                debatecount = list_MTBECE.count(listDec + list_TBECE_rebate)
                ws_rebate.cell(j, 17, debatecount)

    j += 1

wb_rebate.save(table_rebate)
