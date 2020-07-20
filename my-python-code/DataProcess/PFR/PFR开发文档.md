# PFR报表开发说明书

## 概述 

在使用前,需维护TeamTemplate.xlsx的如下sheet 

**2.DESA&PFR-LOB** 用于生成LOB

**6.PFR-EA EBU**  用于生成EA EBU ;

**7.PFR-Exchange Rate** 用于生成汇率 ;

**8.PFR-FCG Name**  用于计算 Discount 和 时 与主表匹配 Customer的对应关系 ;

**9.PFR-Other Sales Adj.** 用于生成Other Sales Adj.列 ;

**10.PFR-Other GM Adj.** 用于生成Other GM Adj.列 ;

**11.PFR-SAR Adj.** 用于生成SAR列 ;



提供rebateAmount.xlsx用于计算 Discount 和  Purchasing rebate 

## 1.生成UnitPrice列

Unit Price = Net Sales / Units

单价 = 净销售额 / 数量



## 2.生成Unit Cost列

单位成本计算分两种情况



### 2.1 内销

MFG Plant 属于 BHO/CQP/DFM/XCE 

Unit Cost =  Material / Units

单位成本 = 原材料 / 数量

### 2.2 进口

Unit Cost = Prod Cost / Units

单位成本 = 生产成本 / 数量



## 3.生成Unit GM列

Unit GM = Unit Price - Unit Cost

单台利润 = 单价 - 单台成本



## 4.生成Unit GM%列

Unit GM% = Unit GM / Unit Price

单台利润率 = 单台利润 / 单价



## 5.生成RC列

Application = “CONSTRUCTION” 的更新为587,其它都更新为497



## 6.生成EA EBU列

该功能需要用到主表与mapping关联

PFRsource.xlsx (**主表**) 与 Team Template 的 “6.PFR-EA EBU”  表 (**mapping表**) 关联

遵循以下四个条件

1.主表的Category 在mapping表的范围内，如果没有找到则属于“Other”

2.主表的Application = mapping表的 Application

3.主表的MBU = mapping表的 MBU

4.主表的Family = mapping 表的Family

**隆工特殊处理：**

隆工是不需要维护到mapping表的，如果满足以下三个条件

1.MFG Plant 属于BHO/CQP/DFM/XCE

2.Application 为 CONSTRUCTION

3.Engine Family 为 LONKING SHANGHAI 

那么获取的EA EBU = Domestic DCEC construction



注意：代码里在比对前生成了全字母大写在比较以防止因为大小写不一致导致匹配不到



## 7.生成Exchange Rate列

根据主表匹配 **7.PFR-Exchange Rate** 里的记录，逐个匹配并更新到 Exchange Rate列



## 8.更新Discount列

前提：需提供rebateAcount.xlsx表的Sales 内容

按照如下四个条件匹配，更新 Discount

条件：

主表：PFRsource.xlsx

辅表：rebateAcount.xlsx

mapping表：TeamTemplate.xlsx 的 sheet **《8.PFR-FCG Name》**  和 **《3.DESA&PFR-Emission》**

1 主表的 Acctg Period  对应 辅表 的 Jan 、 Feb 、Mar 等月份

2 对应Customer 

​	step1 主表关联辅表获得 FCG Name -------- Customer 

​	step2 根据该关系 匹配 辅表的 Customer

3 主表的Engine Family = 辅表的 Engine Family

4 对应Emission

​	step1  主表关联mapping表 **3.DESA&PFR-Emission**  获得  Config --------- Emission 对应关系

​	step2  通过这个对应关系关联辅表的 Emission



## 9.生成Purchasing rebate

逻辑同8.更新Discount

前提：需提供rebateAcount.xlsx表的Purchase 内容

按照如下四个条件匹配，更新 Purchasing rebate

条件：

主表：PFRsource.xlsx

辅表：rebateAcount.xlsx

mapping表：TeamTemplate.xlsx 的 sheet **《8.PFR-FCG Name》**  和 **《3.DESA&PFR-Emission》**

1 主表的 Acctg Period  对应 辅表 的 Jan 、 Feb 、Mar 等月份

2 对应Customer 

​	step1 主表关联辅表获得 FCG Name -------- Customer 

​	step2 根据该关系 匹配 辅表的 Customer

3 主表的Engine Family = 辅表的 Engine Family

4 对应Emission

​	step1  主表关联mapping表 **3.DESA&PFR-Emission**  获得  Config --------- Emission 对应关系

​	step2  通过这个对应关系关联辅表的 Emission



注意：主表的月份格式为 “202001”  程序是截取 后两位 “01”作为月份，再与辅表的“Jan”进行匹配，所以系统不判断跨年份的情况

## 10.生成LOB

前提：需维护 《2.DESA&PFR-LOB》

程序根据 EA EBU 与 LOB 的对应关系自动更新LOB



## 11.生成Other Sales Adj.列

前提：需维护《9.PFR-Other Sales Adj.》mapping表

逻辑：相同LOB按照月份汇总Units

单个Units / 汇总的Units  *  分摊系数

其中分摊系数是从mapping匹配到



实现逻辑：

step1 定义字典 dict_target_lob_month_units 以 list(LOB,Month)作为Key，以Units作为Value,然后字典里全部赋值为0。

```Python
# 获取LOB
target_lob = ws_target['AR':'AR']
target_lob_list = [str(e.value) for e in target_lob]

# 获取Month
target_month = ws_target['D':'D']
target_month_list = [str(e.value)[-2:] for e in target_month]

# 获取units的值
target_units = ws_target['U':'U']
target_units_list = [str(e.value) for e in target_units]

# 定义字典 dict_target_lob_month_units 以 list(LOB,Month)作为Key，以Units作为Value
dict_target_lob_month_units = dict(zip(zip(target_lob_list, target_month_list), target_units_list))

# 遍历字典赋值为0
for key in dict_target_lob_month_units:
    dict_target_lob_month_units[key] = 0

```

step2：获得所有LOB-月份 的units 汇总值，遍历全表逐个判断是否存在 LOB-Month 为键，如果存在则累加。

```python
dict_tr = {}
# 12.1 遍历求得 每个 LOB 在各个月份的 汇总值 ， 如 LCV 1月份 汇总 2月份汇总。。。
tr = 2
while tr <= ws_target.max_row:
    # 获取LOB
    tr_lob = ws_target.cell(tr, 44).value
    # 获取Month
    tr_month = str(ws_target.cell(tr, 4).value)[-2:]
    # 获取units的值
    tr_units = str(ws_target.cell(tr, 21).value)

    dict_tr = {(tr_lob, tr_month): tr_units}
    # dict_target_lob_month_units  units 的总和
    # 判断是否存在 LOB-Month 为键，如果存在则累加
    if (tr_lob, tr_month) in dict_target_lob_month_units:
        dict_target_lob_month_units[(tr_lob, tr_month)] = 		     dict_target_lob_month_units[(tr_lob, tr_month)] + int(
            tr_units)

    tr += 1
```

step3：获取 LOB-Month 的调整系数 Other Sales Adj 

以LOB和Month 为Key , 调整系数为Value，其中Month是 “01”拼出来的

```Python
# 12.2 从模板表取出 Other Sales Adj
osa = 2
dict_osa_jan = {}
dict_osa_feb = {}
dict_osa_mar = {}
dict_osa_apr = {}
dict_osa_may = {}
dict_osa_jun = {}
dict_osa_jul = {}
dict_osa_aug = {}
dict_osa_sep = {}
dict_osa_oct = {}
dict_osa_nov = {}
dict_osa_dec = {}
while osa <= ws_template_Other_Sales_Adj.max_row:
    osa_lob = str(ws_template_Other_Sales_Adj.cell(osa, 1).value)
    osa_sales_jan = str(ws_template_Other_Sales_Adj.cell(osa, 2).value)
    osa_sales_feb = str(ws_template_Other_Sales_Adj.cell(osa, 3).value)
    osa_sales_mar = str(ws_template_Other_Sales_Adj.cell(osa, 4).value)
    osa_sales_apr = str(ws_template_Other_Sales_Adj.cell(osa, 5).value)
    osa_sales_may = str(ws_template_Other_Sales_Adj.cell(osa, 6).value)
    osa_sales_jun = str(ws_template_Other_Sales_Adj.cell(osa, 7).value)
    osa_sales_jul = str(ws_template_Other_Sales_Adj.cell(osa, 8).value)
    osa_sales_aug = str(ws_template_Other_Sales_Adj.cell(osa, 9).value)
    osa_sales_sep = str(ws_template_Other_Sales_Adj.cell(osa, 10).value)
    osa_sales_oct = str(ws_template_Other_Sales_Adj.cell(osa, 11).value)
    osa_sales_nov = str(ws_template_Other_Sales_Adj.cell(osa, 12).value)
    osa_sales_dec = str(ws_template_Other_Sales_Adj.cell(osa, 13).value)

    # 拼出Key
    dict_osa_jan[(osa_lob, "01")] = osa_sales_jan
    dict_osa_feb[(osa_lob, "02")] = osa_sales_feb
    dict_osa_mar[(osa_lob, "03")] = osa_sales_mar
    dict_osa_apr[(osa_lob, "04")] = osa_sales_apr
    dict_osa_may[(osa_lob, "05")] = osa_sales_may
    dict_osa_jun[(osa_lob, "06")] = osa_sales_jun
    dict_osa_jul[(osa_lob, "07")] = osa_sales_jul
    dict_osa_aug[(osa_lob, "08")] = osa_sales_aug
    dict_osa_sep[(osa_lob, "09")] = osa_sales_sep
    dict_osa_oct[(osa_lob, "10")] = osa_sales_oct
    dict_osa_nov[(osa_lob, "11")] = osa_sales_nov
    dict_osa_dec[(osa_lob, "12")] = osa_sales_dec

    # 获得 Other Sales Adj. 每个LOB的每个月份的值 用于后续分摊
    dict_osa = {**dict_osa_jan, **dict_osa_feb, **dict_osa_mar, **dict_osa_apr, **dict_osa_may, **dict_osa_jun,
                **dict_osa_jul, **dict_osa_aug, **dict_osa_sep, **dict_osa_oct, **dict_osa_nov, **dict_osa_dec}

    osa += 1
```

step4：遍历Target结果表逐行按权重分摊调整系数

理解下面三个字典

**dict_tr**  Target表 单行 LOB,Month :  Units

**dict_target_lob_month_units**  Target表的 累加值 LOB,Month :  Units

**dict_osa** 维护表的系数 LOB,Month :  cell.value

逐行判断以 LOB，Month 为Key 是否存在于 dict_target_lob_month_units

如果存在 单行Units / 汇总Units * 调整系数

```Python
osa_w = 2
while osa_w <= ws_target.max_row:
    # 获取LOB
    w_lob = ws_target.cell(osa_w, 44).value
    # 获取Month
    w_month = str(ws_target.cell(osa_w, 4).value)[-2:]
    # 获取units的值
    w_units = str(ws_target.cell(osa_w, 21).value)

    dict_tr = {(w_lob, w_month): w_units}
    if (w_lob, w_month) in dict_target_lob_month_units:
        if ( dict_tr[(w_lob, w_month)] == 'None' or dict_target_lob_month_units[(w_lob, w_month)] == 'None' or dict_osa[(w_lob, w_month)] == 'None' ):
            w_osa =0
        else:
            w_osa = int(dict_tr[(w_lob, w_month)]) / int(dict_target_lob_month_units[(w_lob, w_month)]) * int(dict_osa[(w_lob, w_month)])
        ws_target.cell(osa_w, 46, w_osa)

    if (w_lob, w_month) in dict_target_lob_month_units:
        if (dict_tr[(w_lob, w_month)] == 'None' or dict_target_lob_month_units[(w_lob, w_month)] == 'None' or
                dict_oga[(w_lob, w_month)] == 'None'):
            w_osa = 0
        else:
            w_osa = int(dict_tr[(w_lob, w_month)]) / int(dict_target_lob_month_units[(w_lob, w_month)]) * int(
                dict_oga[(w_lob, w_month)])
        ws_target.cell(osa_w, 47, w_osa)

    if (w_lob, w_month) in dict_target_lob_month_units:
        if (dict_tr[(w_lob, w_month)] == 'None' or dict_target_lob_month_units[(w_lob, w_month)] == 'None' or
                dict_sar[(w_lob, w_month)] == 'None'):
            w_sar = 0
        else:
            w_sar = int(dict_tr[(w_lob, w_month)]) / int(dict_target_lob_month_units[(w_lob, w_month)]) * int(
                dict_sar[(w_lob, w_month)])
        ws_target.cell(osa_w, 48, w_sar)

    osa_w += 1
```



## 12.生成Other GM Adj.列

前提：需维护《10.PFR-Other GM Adj.》mapping表

逻辑：同11生成Other Sales Adj.列



## 13.生成SAR列

前提：需维护《11.PFR-SAR Adj.》mapping表

逻辑：同11生成Other Sales Adj.列



## 14 生成缩减版报告

**需求：**PFR的数据源不变，需要生成的结果表上删减掉以下十五列

A列 ID

F列  Engine Family Code

H列 Engine Application Code

E列  ABO Name

M列 ABO Group Name

N列 WPT Name

O列 Trans FCG

Q列 Trans FCG Name

R列 Forecaster Id

S列 Forecaster Name

T列 Grs Op Cont

Z列 % Op Cont

AA列 Version

AH列 Scenario

AI列 Tlva Indic

**结果：**

程序为“PFRTarget.xlsx”生成两个sheet表

1. reduce version : 为删减后的结果表

2. full version :为完整版的结果表

**实现:**

以生成full version为基础，生成reduce version，并把不需要删除的列逐行写入到reduce version。