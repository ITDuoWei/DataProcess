import openpyxl
import datetime

starttime = datetime.datetime.now()


doc = """

  _____       _   _                   _____                   _             
 |  __ \     | | | |                 |  __ \                 (_)            
 | |__) |   _| |_| |__   ___  _ __   | |__) |   _ _ __  _ __  _ _ __   __ _ 
 |  ___/ | | | __| '_ \ / _ \| '_ \  |  _  / | | | '_ \| '_ \| | '_ \ / _` |
 | |   | |_| | |_| | | | (_) | | | | | | \ \ |_| | | | | | | | | | | | (_| |
 |_|    \__, |\__|_| |_|\___/|_| |_| |_|  \_\__,_|_| |_|_| |_|_|_| |_|\__, |
         __/ |                                                         __/ |
        |___/                                                         |___/
--------------Please have a cup of coffee and wait patiently--------------- 
"""
print(doc)

PFR_source = "PFRsource.xlsx"
PFR_target = "PFRTarget.xlsx"
Tem_source = "PFRTeamTemplate.xlsx"

wb_source = openpyxl.load_workbook(PFR_source)
ws_source = wb_source.worksheets[0]

# 目标表
wb_target = openpyxl.Workbook()
ws_target = wb_target.worksheets[0]

ls_table_source = list(ws_source.values)
ls_table_values_source = [t for t in ls_table_source]

for r in ls_table_source:
    ws_target.append(r)
wb_target.save(PFR_target)

#  模板表
wb_template = openpyxl.load_workbook(Tem_source)
ws_template = wb_template.worksheets[0]

r = 2
while r <= ws_target.max_row:

    ws_source_MFGPlant = str(ws_source.cell(r, 5).value).upper()
    ws_source_EngineFamily = str(ws_source.cell(r, 7).value).upper()
    ws_source_Application = str(ws_source.cell(r, 9).value).upper()
    ws_source_MBUName = str(ws_source.cell(r, 11).value).upper()
    ws_source_FCGName = str(ws_source.cell(r, 16).value).upper()

    ws_source_Units = ws_source.cell(r, 21).value
    ws_source_NetSales = ws_source.cell(r, 24).value
    ws_source_ProdCost = ws_source.cell(r, 25).value
    ws_source_Material = ws_source.cell(r, 30).value
    ws_source_Conversion = ws_source.cell(r, 31).value

    try:
        UnitPrice = ws_source_NetSales / ws_source_Units
        ws_target.cell(r, 38, UnitPrice)
    except:
        ws_target.cell(r, 38, "")

    # 内销
    if "BHO/CQP/DFM/XCE".find(ws_source_MFGPlant) >= 0:
        try:
            UnitCost = ws_source_Material / ws_source_Units
            ws_target.cell(r, 39, UnitCost)
        except:
            ws_target.cell(r, 39, "")
    else:  # 进口
        try:
            UnitCost = ws_source_ProdCost / ws_source_Units
            ws_target.cell(r, 39, UnitCost)
        except:
            ws_target.cell(r, 39, "")

    try:
        UnitGM = UnitPrice - UnitCost
        ws_target.cell(r, 40, UnitGM)
    except:
        ws_target.cell(r, 40, "")

    try:
        UnitGM_precent = UnitGM / UnitPrice * 100
        ws_target.cell(r, 41, UnitGM_precent)
    except:
        ws_target.cell(r, 41, "")

    # RC #  ws_source_Application = “CONSTRUCTION” 为 587 其他都是 497
    if ws_source_Application == "CONSTRUCTION":
        ws_target.cell(r, 42, 587)
    else:
        ws_target.cell(r, 42, 497)

    # Team 生成
    t = 2
    while t <= ws_template.max_row:
        ws_template_Category = str(ws_template.cell(t, 1).value).upper()
        ws_template_Family = str(ws_template.cell(t, 4).value).upper()
        ws_template_Application = str(ws_template.cell(t, 2).value).upper()
        ws_template_MBU = str(ws_template.cell(t, 3).value).upper()
        ws_template_EAEBU = str(ws_template.cell(t, 5).value).upper()  # Team

        # other 处理
        if ws_template_Category.find(ws_source_MFGPlant) < 0:
            ws_source_MFGPlant = "OTHER"

        if ws_template_Category.find(ws_source_MFGPlant) >= 0 \
                and ws_source_Application == ws_template_Application \
                and ws_source_MBUName == ws_template_MBU \
                and (ws_source_EngineFamily == ws_template_Family or ws_template_Family is None):
            # print( ws_target.cell(r,37) )
            ws_target.cell(r, 37, ws_template_EAEBU)

        t += 1

        #  隆工特殊处理 Category 属于 BHO/CQP/DFM/XCE  and Application == CONSTRUCTION and Family
        #  隆工不需要维护
        if "BHO/CQP/DFM/XCE".find(ws_source_MFGPlant) >= 0 \
                and ws_source_Application == "CONSTRUCTION" \
                and ws_source_EngineFamily == "B6.7" \
                and ws_source_FCGName == "LONKING SHANGHAI":
            ws_target.cell(r, 37, "Domestic DCEC construction")


    if r % 100 == 0:
        print("Passed " + str(r) + " records,surplus " + str( int(ws_target.max_row - (r / 100)*100)  ) +" records")

    r += 1

wb_target.save(PFR_target)

endtime = datetime.datetime.now()
print("Done! Use seconds " + str((endtime - starttime).seconds))
print("Mission accomplished!Please exit")
input()