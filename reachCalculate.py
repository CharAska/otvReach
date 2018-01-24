# -*-coding:utf-8-*-
# !/usr/bin/python
import xlrd
import xlwt
from tkinter import *
from tkinter import messagebox #消息框模块
import tkinter.filedialog #文件导入模块
import re #正则表达式
import time, sys #合法性验证

#判断使用期限是否超时
if time.strftime('%Y-%m-%d',time.localtime(time.time())) > "2018-12-31":
    sys.exit()

def vbaDownload():
    vba_exl = xlrd.open_workbook("vba.xlsm")
    file_opt = options = {}
    options['defaultextension'] = '.xlsm'
    options['filetypes'] = [('out', '.xlsm')]
    options['initialfile'] = 'out.xlsm'
    options['parent'] = ui_top
    options['title'] = '另存为计算结果'
    vba_exl.save(tkinter.filedialog.asksaveasfilename(**file_opt))

def reachDownload():
    pass

def helpDownload():
    pass

def chooseFile():
    global data
    exlFile = tkinter.filedialog.askopenfilename()
    data = xlrd.open_workbook(exlFile)
    label_fileName.config(text=re.sub(r".*\/", "", exlFile, count=0))

def mod_mix():
    label_pc = Label(ui_top, text="", height=3, width=20)
    label_pc.grid(row=4, column=2)
    label_mob = Label(ui_top, text="", width=20)
    label_mob.grid(row=5, column=2)
    label_ott = Label(ui_top, text="", width=20)
    label_ott.grid(row=6, column=2)
    label_buff = Label(ui_top, text="", width=20)
    label_buff.grid(row=7, column=2)

    label_pc = Label(ui_top, text="", height=3, width=20)
    label_pc.grid(row=4, column=3)
    label_mob = Label(ui_top, text="", width=20)
    label_mob.grid(row=5, column=3)
    label_ott = Label(ui_top, text="", width=20)
    label_ott.grid(row=6, column=3)
    label_buff = Label(ui_top, text="", width=20)
    label_buff.grid(row=7, column=3)

def mod_scale():
    label_pc = Label(ui_top, text="PC曝光占比", height=3, width=20)
    label_pc.grid(row=4, column=2)
    label_mob = Label(ui_top, text="Mobile曝光占比", width=20)
    label_mob.grid(row=5, column=2)
    label_ott = Label(ui_top, text="OTT曝光占比", width=20)
    label_ott.grid(row=6, column=2)
    label_buff = Label(ui_top, text="设备比例浮动系数", width=20)
    label_buff.grid(row=7, column=2)

    # e1 = StringVar()
    global e1,e2,e3,e4
    e1 = Variable()
    e2 = Variable()
    e3 = Variable()
    e4 = Variable()
    e1.set("填写正整数，如0-100")
    e2.set("填写正整数，如0-100")
    e3.set("填写正整数，如0-100")
    e4.set("填写小数，如0.0-1.0")
    entry_pc = Entry(ui_top, textvariable=e1, width=20)
    entry_pc.grid(row=4, column=3)
    entry_mob = Entry(ui_top, textvariable=e2, width=20)
    entry_mob.grid(row=5, column=3)
    entry_ott = Entry(ui_top, textvariable=e3, width=20)
    entry_ott.grid(row=6, column=3)
    entry_buff = Entry(ui_top, textvariable=e4, width=20)
    entry_buff.grid(row=7, column=3)


def exlRead():
    # 完成后需要修改地址，可改成导入方式
    # 曲线表
    global curve_rows, curve_cols, target_rows, target_cols, table_reachCurve, table_target, data
    #data = xlrd.open_workbook('D:\CODE\Python\otvReach\otvReach.xlsx')  # 获取工作表
    table_reachCurve = data.sheet_by_name(u'reachCurve')  # 获取sheet
    curve_rows = table_reachCurve.nrows  # 获取行数
    curve_cols = table_reachCurve.ncols  # 获取列数
    # 目标表
    table_target = data.sheet_by_name(u'Target')  # 获取sheet
    target_rows = table_target.nrows  # 获取行数
    target_cols = table_target.ncols  # 获取列数
    print(curve_rows)
    print(curve_cols)


def curveList():
    # 获取生成curve的List
    global reachCurve
    reachCurve = []
    for i in range(curve_rows):
        row = []
        for j in range(curve_cols):
            row.append(table_reachCurve.cell(i, j).value)
        reachCurve.append(row)


def targetDict():
    # 获取生成target的Dictionary和List
    # 城市, 排期Imp, 排期PC-Imp, 排期Mobile-Imp, 排期OTT-Imp
    # 目标1+%, 目标3+%, 实际1+%, 实际3+%, 目标1+%所需曝光, 目标3+%所需曝光
    global cityDict, cityList, titleList
    cityDict = {}
    for i in range(target_rows):
        cityDict[table_target.cell(i, 0).value] = {}
        for j in range(target_cols):
            cityDict[table_target.cell(i, 0).value][table_target.cell(0, j).value] = table_target.cell(i, j).value
    cityList = []
    for i in range(1, target_rows):
        cityList.append(table_target.cell(i, 0).value)
    titleList = list(cityDict["城市"].keys())

'''
def impTo3plus():
    # 根据排期Imp计算实际3+%
    for city in cityList:
        for i in range(1, curve_rows):
            if reachCurve[i][0] == city and reachCurve[i + 1][0] == city:  # 2行均为目标城市
                if reachCurve[i][1] == cityDict[city]["排期Imp"]:
                    cityDict[city]["实际3+%"] = reachCurve[i][6]
                elif reachCurve[i][1] < cityDict[city]["排期Imp"] and reachCurve[i + 1][1] > cityDict[city]["排期Imp"]:
                    cityDict[city]["实际3+%"] = (cityDict[city]["排期Imp"] - reachCurve[i][1]) * (
                    reachCurve[i + 1][6] - reachCurve[i][6]) \
                                              / (reachCurve[i + 1][1] - reachCurve[i][1]) + reachCurve[i][6]

'''
def aTob(a1, a2, b1, b2):
    # 根据指定目标内容推到实际情况;a:目标内容(int)，b：实际情况(str）;求b2
    for city in cityList:
        for i in range(1, curve_rows - 1):
            if reachCurve[i][0] == city and reachCurve[i + 1][0] == city:  # 2行均为目标城市
                if reachCurve[i + 1][a1] == cityDict[city][b1]:
                    cityDict[city][b2] = reachCurve[i + 1][a2]
                    break
                elif reachCurve[i][a1] < cityDict[city][b1] and reachCurve[i + 1][a1] > cityDict[city][b1]:
                    cityDict[city][b2] = (cityDict[city][b1] - reachCurve[i][a1]) * (
                    reachCurve[i + 1][a2] - reachCurve[i][a2]) \
                                         / (reachCurve[i + 1][a1] - reachCurve[i][a1]) + reachCurve[i][a2]
                    break


def aTob_scale(a1, a2, b1, b2, pc_mob, ott_mob, buff):
    # 根据指定目标内容推到实际情况;a:目标内容(int)，b：实际情况(str）;求b2
    # mob_pc: mob除以pc的系数（float); mob_ott：mob除以ott的系数(float);buff(float,ex:0.2,表示±20%)
    # 对曲线列表进行数据排除,pc2,mobile3,ott4
    global reachCurve_filter
    reachCurve_filter = []
    global curve_rows_filter
    reachCurve_filter.append(reachCurve[0])  # 添加title
    for i in range(1, curve_rows):
        # 筛选比例曲线
        if reachCurve[i][3] != 0:
            if reachCurve[i][2] / reachCurve[i][3] <= pc_mob * (1 + buff) and reachCurve[i][2] / reachCurve[i][
                3] >= pc_mob * (1 - buff) \
                    and reachCurve[i][4] / reachCurve[i][3] <= ott_mob * (1 + buff) and reachCurve[i][4] / \
                    reachCurve[i][3] >= ott_mob * (1 - buff):
                reachCurve_filter.append(reachCurve[i])
    curve_rows_filter = len(reachCurve_filter)  # 获取筛选表行数
    print(curve_rows_filter)  # 筛选的行数
    for city in cityList:
        for i in range(1, curve_rows_filter - 1):
            if reachCurve_filter[i][0] == city and reachCurve_filter[i + 1][0] == city:  # 2行均为目标城市
                if reachCurve_filter[i + 1][a1] == cityDict[city][b1]:
                    cityDict[city][b2] = reachCurve_filter[i + 1][a2]
                    cityDict[city]["所需PC-Imp"] = reachCurve_filter[i + 1][2]  # pc
                    cityDict[city]["所需Mobile-Imp"] = reachCurve_filter[i + 1][3]  # mob
                    cityDict[city]["所需OTT-Imp"] = reachCurve_filter[i + 1][4]  # ott
                    break
                elif reachCurve_filter[i][a1] < cityDict[city][b1] and reachCurve_filter[i + 1][a1] > cityDict[city][
                    b1]:
                    cityDict[city][b2] = (cityDict[city][b1] - reachCurve_filter[i][a1]) * (
                    reachCurve_filter[i + 1][a2] - reachCurve_filter[i][a2]) \
                                         / (reachCurve_filter[i + 1][a1] - reachCurve_filter[i][a1]) + \
                                         reachCurve_filter[i][a2]
                    # PC-imp
                    cityDict[city]["所需PC-Imp"] = (cityDict[city][b1] - reachCurve_filter[i][a1]) * (
                    reachCurve_filter[i + 1][2] - reachCurve_filter[i][2]) \
                                                 / (reachCurve_filter[i + 1][a1] - reachCurve_filter[i][a1]) + \
                                                 reachCurve_filter[i][2]
                    # Mobile-imp
                    cityDict[city]["所需Mobile-Imp"] = (cityDict[city][b1] - reachCurve_filter[i][a1]) * (
                    reachCurve_filter[i + 1][3] - reachCurve_filter[i][3]) \
                                                     / (reachCurve_filter[i + 1][a1] - reachCurve_filter[i][a1]) + \
                                                     reachCurve_filter[i][3]
                    # OTT-imp
                    cityDict[city]["所需OTT-Imp"] = (cityDict[city][b1] - reachCurve_filter[i][a1]) * (
                    reachCurve_filter[i + 1][4] - reachCurve_filter[i][4]) \
                                                  / (reachCurve_filter[i + 1][a1] - reachCurve_filter[i][a1]) + \
                                                  reachCurve_filter[i][4]
                    break

def aTob_find(a1, a2, b1, b2):
    # 根据指定目标内容推到实际情况;a:目标内容(int)，b：实际情况(str）;求b2
    for city in cityList:
        for i in range(1, curve_rows - 1):
            if reachCurve[i][0] == city and reachCurve[i + 1][0] == city:  # 2行均为目标城市
                if reachCurve[i + 1][a1] == cityDict[city][b1]:
                    cityDict[city][b2] = reachCurve[i + 1][a2]
                    break
                elif reachCurve[i][a1] < cityDict[city][b1] and reachCurve[i + 1][a1] > cityDict[city][b1]:
                    cityDict[city][b2] = reachCurve[i + 1][a2]
                    break


def aTob_scale_find(a1, a2, b1, b2, pc_mob, ott_mob, buff):
    # 根据指定目标内容推到实际情况;a:目标内容(int)，b：实际情况(str）;求b2
    # mob_pc: mob除以pc的系数（float); mob_ott：mob除以ott的系数(float);buff(float,ex:0.2,表示±20%)
    # 对曲线列表进行数据排除,pc2,mobile3,ott4
    global reachCurve_filter
    reachCurve_filter = []
    global curve_rows_filter
    reachCurve_filter.append(reachCurve[0])  # 添加title
    for i in range(1, curve_rows):
        # 筛选比例曲线
        if reachCurve[i][3] != 0:
            if reachCurve[i][2] / reachCurve[i][3] <= pc_mob * (1 + buff) and reachCurve[i][2] / reachCurve[i][
                3] >= pc_mob * (1 - buff) \
                    and reachCurve[i][4] / reachCurve[i][3] <= ott_mob * (1 + buff) and reachCurve[i][4] / \
                    reachCurve[i][3] >= ott_mob * (1 - buff):
                reachCurve_filter.append(reachCurve[i])
    curve_rows_filter = len(reachCurve_filter)  # 获取筛选表行数
    #print(curve_rows_filter)  # 筛选的行数
    for city in cityList:
        for i in range(1, curve_rows_filter - 1):
            if reachCurve_filter[i][0] == city and reachCurve_filter[i + 1][0] == city:  # 2行均为目标城市
                if reachCurve_filter[i + 1][a1] == cityDict[city][b1]:
                    cityDict[city][b2] = reachCurve_filter[i + 1][a2]
                    cityDict[city]["所需PC-Imp"] = reachCurve_filter[i + 1][2]  # pc
                    cityDict[city]["所需Mobile-Imp"] = reachCurve_filter[i + 1][3]  # mob
                    cityDict[city]["所需OTT-Imp"] = reachCurve_filter[i + 1][4]  # ott
                    break
                elif reachCurve_filter[i][a1] < cityDict[city][b1] and reachCurve_filter[i + 1][a1] > cityDict[city][b1]:
                    cityDict[city][b2] = reachCurve_filter[i + 1][a2]
                    cityDict[city]["所需PC-Imp"] = reachCurve_filter[i + 1][2]  # pc
                    cityDict[city]["所需Mobile-Imp"] = reachCurve_filter[i + 1][3]  # mob
                    cityDict[city]["所需OTT-Imp"] = reachCurve_filter[i + 1][4]  # ott
                    break



def exlWrite():
    # 创建工作簿
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建sheet
    out_sheet = workbook.add_sheet('demo', cell_overwrite_ok=True)
    # out_sheet.write(0, 0, "abc")
    for i in range(target_cols):
        out_sheet.write(0, i, titleList[i])
    for j in range(1, target_rows):
        out_sheet.write(j, 0, cityList[j - 1])
    for i in range(1, target_rows):
        for j in range(target_cols):
            out_sheet.write(i, j, cityDict[cityList[i - 1]][titleList[j]])
    # 保存文件
    #workbook.save('D:\CODE\Python\otvReach\out.xls')
    file_opt = options = {}
    options['defaultextension'] = '.xls'
    options['filetypes'] = [('out', '.xls')]
    options['initialfile'] = 'out.xls'
    options['parent'] = ui_top
    options['title'] = '另存为计算结果'
    workbook.save(tkinter.filedialog.asksaveasfilename(**file_opt))

########################################################################################################################
########################################################################################################################
# 运行
def upload():
    exlRead()
    curveList()
    targetDict()
    messagebox.showinfo(message="导入完成")

def checkData(b1):
    for city in cityList:
        if cityDict[city][b1] == "":
            messagebox.showinfo(message="请确认"+b1+"已填写数据")
            return 0

def calculateReach():
    # 界面tkinter变量 device_radio ; ck1-4 ; e1-4;需要配合float()使用
    #messagebox.showinfo(message= float(e1.get())+float(e2.get()))
    # aTob_scale(7, 1, "目标3+%", "目标3+%所需曝光", 2 / 7, 1 / 7, 0.2)
    global workWell
    workWell = 0
    if device_radio.get() == 1:
        try:
            if int(e1.get()) + int(e2.get()) + int(e3.get()) <= 0 or int(e1.get()) < 0 or int(e2.get()) < 0 or int(
                    e3.get()) < 0:
                messagebox.showinfo(message="占比请填写正数，并保证总和大于0")
                return 0
        except:
            messagebox.showinfo(message="占比请填写正整数")
            return 0
        try:
            if float(e4.get()) < 0:
                messagebox.showinfo(message="浮动系数请填写正数")
                return 0
        except:
            messagebox.showinfo(message="浮动系数请填写正数")
            return 0
        if ck1.get() == 1:#曝光求1+%
            checkData( "排期Imp")
            aTob_scale(1, 5, "排期Imp", "实际1+%", float(e1.get()) / float(e2.get()), float(e3.get()) / float(e2.get()), float(e4.get()))
            workWell = 1
        if ck2.get() == 1:
            checkData("排期Imp")
            aTob_scale(1, 7, "排期Imp", "实际3+%", float(e1.get()) / float(e2.get()), float(e3.get()) / float(e2.get()), float(e4.get()))
            workWell = 1
        if ck3.get() == 1:
            checkData("目标1+%")
            aTob_scale(5, 1, "目标1+%", "目标1+%所需曝光", float(e1.get()) / float(e2.get()), float(e3.get()) / float(e2.get()), float(e4.get()))
            workWell = 1
        if ck4.get() == 1:
            checkData( "目标3+%")
            aTob_scale(7, 1, "目标3+%", "目标3+%所需曝光", float(e1.get()) / float(e2.get()), float(e3.get()) / float(e2.get()), float(e4.get()))
            workWell = 1
    elif device_radio.get() == 2:
        if ck1.get() == 1:#曝光求1+%
            checkData("排期Imp")
            aTob(1, 5, "排期Imp", "实际1+%")
            workWell = 1
        if ck2.get() == 1:
            checkData("排期Imp")
            aTob(1, 7, "排期Imp", "实际3+%")
            workWell = 1
        if ck3.get() == 1:
            checkData( "目标1+%")
            aTob(5, 1, "目标1+%", "目标1+%所需曝光")
            workWell = 1
        if ck4.get() == 1:
            checkData( "目标3+%")
            aTob(7, 1, "目标3+%", "目标3+%所需曝光")
            workWell = 1

    if workWell == 1:
        messagebox.showinfo(message="计算完成，请保存结果")
        exlWrite()
    else:
        messagebox.showinfo(message="计算未完成，请确保文件准确")

def findReach():
        # 界面tkinter变量 device_radio ; ck1-4 ; e1-4;需要配合float()使用
        # messagebox.showinfo(message= float(e1.get())+float(e2.get()))
        # aTob_scale(7, 1, "目标3+%", "目标3+%所需曝光", 2 / 7, 1 / 7, 0.2)
        global workWell
        workWell = 0
        if device_radio.get() == 1:
            try:
                if int(e1.get()) + int(e2.get()) + int(e3.get()) <= 0 or int(e1.get()) < 0 or int(e2.get()) < 0 or int(
                        e3.get()) < 0:
                    messagebox.showinfo(message="占比请填写正数，并保证总和大于0")
                    return 0
            except:
                messagebox.showinfo(message="占比请填写正整数")
                return 0
            try:
                if float(e4.get()) < 0:
                    messagebox.showinfo(message="浮动系数请填写正数")
                    return 0
            except:
                messagebox.showinfo(message="浮动系数请填写正数")
                return 0
            if ck1.get() == 1:  # 曝光求1+%
                checkData("排期Imp")
                aTob_scale_find(1, 5, "排期Imp", "实际1+%", float(e1.get()) / float(e2.get()), float(e3.get()) / float(e2.get()),
                           float(e4.get()))
                workWell = 1
            if ck2.get() == 1:
                checkData("排期Imp")
                aTob_scale_find(1, 7, "排期Imp", "实际3+%", float(e1.get()) / float(e2.get()), float(e3.get()) / float(e2.get()),
                           float(e4.get()))
                workWell = 1
            if ck3.get() == 1:
                checkData("目标1+%")
                aTob_scale_find(5, 1, "目标1+%", "目标1+%所需曝光", float(e1.get()) / float(e2.get()),
                           float(e3.get()) / float(e2.get()), float(e4.get()))
                workWell = 1
            if ck4.get() == 1:
                checkData("目标3+%")
                aTob_scale_find(7, 1, "目标3+%", "目标3+%所需曝光", float(e1.get()) / float(e2.get()),
                           float(e3.get()) / float(e2.get()), float(e4.get()))
                workWell = 1
        elif device_radio.get() == 2:
            if ck1.get() == 1:  # 曝光求1+%
                checkData("排期Imp")
                aTob_find(1, 5, "排期Imp", "实际1+%")
                workWell = 1
            if ck2.get() == 1:
                checkData("排期Imp")
                aTob_find(1, 7, "排期Imp", "实际3+%")
                workWell = 1
            if ck3.get() == 1:
                checkData("目标1+%")
                aTob_find(5, 1, "目标1+%", "目标1+%所需曝光")
                workWell = 1
            if ck4.get() == 1:
                checkData("目标3+%")
                aTob_find(7, 1, "目标3+%", "目标3+%所需曝光")
                workWell = 1

        if workWell == 1:
            messagebox.showinfo(message="计算完成，请保存结果")
            exlWrite()
        else:
            messagebox.showinfo(message="计算未完成，请确保文件准确")

# test code
# exlRead()
# curveList()
# targetDict()

# aTob(1,5,"排期Imp","实际1+%")#混合曲线计算1+
# aTob(1,7,"排期Imp","实际3+%")#混合曲线计算3+
# aTob(5,1,"目标1+%","目标1+%所需曝光")#混合曲线计算1+%所需曝光
# aTob(7,1,"目标3+%","目标3+%所需曝光")#混合曲线计算3+%所需曝光
# aTob_scale(7,1,"目标3+%","目标3+%所需曝光",2/7,1/7,0.2)
# exlWrite()
# print (titleList)
# print (cityDict["上海"])
# print (cityDict["北京"])

########################################################################################################################
########################################################################################################################
# GUI界面代码
ui_top = Tk()
#ui_top.title('OTV项目Reach预估工具 Ver0.1')
# 边框
label_top = Label(ui_top, text="", width=3)
label_top.grid(row=0, column=0)
#label_sign = Label(ui_top, text="Ver0.1 by Char", height=3, width=20, anchor="sw")
#label_sign.grid(row=8, column=1)
label_end = Label(ui_top, text="", width=3)
label_end.grid(row=9, column=4)
'''
# 第1行-导入excel并生成计算list及dict
button_vbaDownload = Button(ui_top, text="曲线制作表下载", width=20, command=lambda: vbaDownload())
button_vbaDownload.grid(row=1, column=1)
button_reachDownload = Button(ui_top, text="Reach计算表下载", width=20, command=lambda: reachDownload())
button_reachDownload.grid(row=1, column=2)
button_helpDownload = Button(ui_top, text="说明文档下载", width=20, command=lambda: helpDownload())
button_helpDownload.grid(row=1, column=3)
'''
# 第2行-导入excel并生成计算list及dict
button_chooseFile = Button(ui_top, text="选择文件", width=20, command=lambda: chooseFile())
button_chooseFile.grid(row=2, column=1)
label_fileName = Label(ui_top, text="尚未上传", height=3, width=20)
label_fileName.grid(row=2, column=2)
button_upload = Button(ui_top, text="导入", width=20, command=lambda: upload())
button_upload.grid(row=2, column=3)
# 第3行-曲线是否设备mix
label_device = Label(ui_top, text="曲线设备是否分离", height=3, width=20)
label_device.grid(row=3, column=1)
device_radio = IntVar()
device_radio.set(1)
radio_1 = Radiobutton(ui_top, text='区分设备', variable=device_radio, value=1, width=20,command=lambda: mod_scale())
radio_2 = Radiobutton(ui_top, text='不区分设备', variable=device_radio, value=2, width=20,command=lambda: mod_mix())
radio_1.grid(row=3, column=2)
radio_2.grid(row=3, column=3)
# 计算方式
ck1 = IntVar()
ck1.set(0)
check_1 = Checkbutton(ui_top, text='已知曝光求1+%', variable=ck1, onvalue=1, offvalue=0, height=3, width=20)
ck2 = IntVar()
ck2.set(0)
check_2 = Checkbutton(ui_top, text='已知曝光求3+%', variable=ck2, onvalue=1, offvalue=0, height=3, width=20)
ck3 = IntVar()
ck3.set(0)
check_3 = Checkbutton(ui_top, text='已知1+%求曝光', variable=ck3, onvalue=1, offvalue=0, height=3, width=20)
ck4 = IntVar()
ck4.set(0)
check_4 = Checkbutton(ui_top, text='已知3+%求曝光', variable=ck4, onvalue=1, offvalue=0, height=3, width=20)
check_1.grid(row=4, column=1)
check_2.grid(row=5, column=1)
check_3.grid(row=6, column=1)
check_4.grid(row=7, column=1)
# 参数设定
'''
label_pc = Label(ui_top, text="", height=3, width=20)
label_pc.grid(row=4, column=2)
label_mob = Label(ui_top, text="", width=20)
label_mob.grid(row=5, column=2)
label_ott = Label(ui_top, text="", width=20)
label_ott.grid(row=6, column=2)
label_buff = Label(ui_top, text="", width=20)
label_buff.grid(row=7, column=2)

label_pc1 = Label(ui_top, text="", height=3, width=20)
label_pc1.grid(row=4, column=3)
label_mob1 = Label(ui_top, text="", width=20)
label_mob1.grid(row=5, column=3)
label_ott1 = Label(ui_top, text="", width=20)
label_ott1.grid(row=6, column=3)
label_buff1 = Label(ui_top, text="", width=20)
label_buff1.grid(row=7, column=3)
'''
label_pc = Label(ui_top, text="PC曝光占比", height=3, width=20)
label_pc.grid(row=4, column=2)
label_mob = Label(ui_top, text="Mobile曝光占比", width=20)
label_mob.grid(row=5, column=2)
label_ott = Label(ui_top, text="OTT曝光占比", width=20)
label_ott.grid(row=6, column=2)
label_buff = Label(ui_top, text="设备比例浮动系数", width=20)
label_buff.grid(row=7, column=2)

# e1 = StringVar()
global e1,e2,e3,e4
e1 = Variable()
e2 = Variable()
e3 = Variable()
e4 = Variable()
e1.set("填写正整数，如0-100")
e2.set("填写正整数，如0-100")
e3.set("填写正整数，如0-100")
e4.set("填写小数，如0.0-1.0")
entry_pc = Entry(ui_top, textvariable=e1, width=20)
entry_pc.grid(row=4, column=3)
entry_mob = Entry(ui_top, textvariable=e2, width=20)
entry_mob.grid(row=5, column=3)
entry_ott = Entry(ui_top, textvariable=e3, width=20)
entry_ott.grid(row=6, column=3)
entry_buff = Entry(ui_top, textvariable=e4, width=20)
entry_buff.grid(row=7, column=3)

# var.get()，判断是否选中，返回0或1
# 运行button
button_find = Button(ui_top, text="查找Reach", height=3, width=20, command=lambda: findReach())
button_find.grid(row=8, column=2)
button_calculate = Button(ui_top, text="预估计算", height=3, width=20, command=lambda: calculateReach())
button_calculate.grid(row=8, column=3)
# 进入消息循环
ui_top.title('OTV项目Reach预估工具 Ver0.2')
label_sign = Label(ui_top, text="Ver0.2 by Char", height=3, width=20, anchor="sw")
label_sign.grid(row=8, column=1)
ui_top.mainloop()
