# -*- coding: utf_8 -*-

import base64
import configparser
import os
import shutil
import sqlite3
import time
import tkinter as tk
from tkinter import messagebox, ttk
import sys

from docxtpl import DocxTemplate, InlineImage
# for height and width you have to use millimeters (Mm), inches or points(Pt) class :
from docx.shared import Mm, Inches, Pt

from icon import img

#python-docx-template

os.chdir(os.path.abspath(os.path.dirname(__file__)))
# print(os.getcwd())
# print(sys.getdefaultencoding())
# print(time.strftime("%Y-%m-%d", time.localtime()))
'''
# 用于将icon.ico图标转化为base64字符串
import base64
with open("icon.ico","rb") as open_icon:
    b64str = base64.b64encode(open_icon.read()) #python3字符都为unicode编码,b64encode函数的参数为byte类型
# write_data = 'img = "%s"' % b64str
write_data = 'img = ' + str(b64str, 'utf-8')  #字符串被b''包围，采用此方法去掉
f = open("icon.py","w")
f.write(write_data)
f.close()
'''


class CaseInfo:
    def __init__(self, vals: list):
        self.id = str(vals[0])
        self.file = str(vals[1])
        self.name = str(vals[2])
        self.age = str(vals[3])
        self.vals = tuple([self.id, self.file, self.name, self.age])


class CaseApp:
    def __init__(self):
        self.appConfig = configparser.ConfigParser()
        self.appConfigFile = 'config/settings/app.txt'
        self.setstr = ('serial', 'name', 'age', 'file', 'number', 'department',
                       'doctor', 'hospitalnum', 'bednum', 'receive', 'handle',
                       'address', 'sampletype', 'samplesize', 'testitem',
                       'samplequality', 'tester', 'collator', 'testresult',
                       'reporttime', 'hospital')
        self.conn = sqlite3.connect('config/database/case.db')
        self.conn.text_factory = str
        self.cursor = self.conn.cursor()
        self.cursor.execute("SELECT tbl_name FROM sqlite_master WHERE type='table'")
        values = self.cursor.fetchall()
        if len(values) == 0 or not 'infoBase' in values[0]:
            self.cursor.execute('''create table if not exists infoBase
                                (id INTEGER primary key autoincrement,
                                serial text,
                                name text,
                                age text,
                                file text,
                                number text,
                                department text,
                                doctor text,
                                hospitalnum text,
                                bednum text,
                                receive text,
                                handle text,
                                address text,
                                sampletype text,
                                samplesize text,
                                testitem text,
                                samplequality text,
                                tester text,
                                collator text,
                                testresult text,
                                reporttime text,
                                hospital text
                                )''')
            messagebox.showerror('错误','数据库丢失，已创建新的数据库，请重新启动程序。')
            # self.cursor.close()
        self.tableItem = 4
        self.selectID = 0 # 选中行数据的ID
        self.selectSerial = 0  # 选中行数据的serial
        self.__infoAttri = ("  ", "编码", "姓名", "年龄", "文件", "电话*",
                            "申请医生", "住院号", "床位号", "接收时间",
                            "处理时间", "报告时间", "地址") # 表格数据
        self.__infoAttriEn = ('id', "serial", 'name', 'age', "file", 'number', 'department',
                              'doctor', 'hospitalnum', 'bednum', 'receive',
                              'handle', 'reporttime', 'address')
        self.__anchor = ('w', 'center', 'w', 'center')
        self.__loadAppConfig()
        self.__setupWidget()

    ########### 界面 ##############
    def __setupWidget(self):
        self.window = tk.Tk()
        self.window.protocol('WM_DELETE_WINDOW', self.__closeWindow)
        self.window.title('Case Helper')
        icon = open('app.ico', 'wb+')
        icon.write(base64.b64decode(img))
        icon.close()
        self.window.iconbitmap("app.ico")
        os.remove("app.ico")
        # self.window.geometry('600x400')
        # w, h = self.window.maxsize()
        w, h = 892, 522
        # 获取屏幕 宽、高
        ws = self.window.winfo_screenwidth()
        hs = self.window.winfo_screenheight()
        # 计算 x, y 位置
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2) - 20
        self.window.geometry('{}x{}+{}+{}'.format(w, h, int(x), int(y)))
        self.window.resizable(0, 0)  # 防止窗口调整大小

        listNameLbl = tk.Label(
            self.window, text='数据列表', font=('微软雅黑', 10), width=15)
        listNameLbl.grid(row=0, column=0)
        tk.Label(self.window, text='基本信息', font=('微软雅黑', 10), width=15).grid(row=0, column=1)

        self.leftFrame = tk.Frame(self.window, width=420, height=h - 30)
        self.rightFrame = tk.Frame(
            self.window, width=470, height=h - 30) #, bg='red')
        self.leftFrame.grid(row=1, column=0, padx=5)
        self.rightFrame.grid(row=1, column=1, sticky=tk.N)

        tk.Label(self.rightFrame, text='检测信息', font=('微软雅黑', 10), width=15).grid(row=6, column=0, columnspan=6)
        # tk.Label(self.rightFrame, text='数据操作', font=('微软雅黑', 10), width=15).grid(row=14, column=0, columnspan=6)

        self.table = ttk.Treeview(
            self.leftFrame,
            show="headings",
            height=23,
            columns=self.__infoAttri[:self.tableItem])
        # self.table.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.table.bind("<Button-1>", self.get_row_value)
        self.vbar = ttk.Scrollbar(
            self.leftFrame, orient='vertical', command=self.table.yview)
        # self.table["columns"]=("姓名","年龄","身高")
        self.table.configure(
            yscrollcommand=self.vbar.set)  #yscrollcommand与Scrollbar的set绑定
        treeItemWidth = int((390-50) / (self.tableItem-1))
        wid = [50, treeItemWidth, treeItemWidth, treeItemWidth]
        index = 0
        for attri in self.__infoAttri[:self.tableItem]:
            self.table.column(
                attri, width=wid[index],
                anchor=self.__anchor[index])  #表示列,不显示
            self.table.heading(attri, text=attri)  #显示表头
            index += 1
        self.table.grid(row=0, column=0, sticky=tk.NS, pady=5)
        self.vbar.grid(row=0, column=1, sticky=tk.NS)
        self.__setupInputs()
        self.__initTable()
        self.currentItem = None
        self.__setupButtons()
        self.window.mainloop()

    def __initTable(self):
        infoList = self.getHospitalCase(self.__curHospital)
        self.updateTable(infoList)

    def __setupInputs(self):
        self.__serialVar = tk.StringVar()
        self.__fileVar = tk.StringVar()
        self.__nameVar = tk.StringVar()
        self.__ageVar = tk.StringVar()
        self.__numberVar = tk.StringVar()
        self.__departmentVar = tk.StringVar()
        self.__doctorVar = tk.StringVar()
        self.__bedNumVar = tk.StringVar()
        self.__hospitalNumVar = tk.StringVar()
        self.__receiveVar = tk.StringVar()
        self.__handleVar = tk.StringVar()
        self.__addressVar = tk.StringVar()

        self.__sampleType = tk.StringVar()
        self.__sampleSize = tk.StringVar()
        self.__testItem = tk.StringVar()
        self.__sampleQuality = tk.StringVar()
        self.__tester = tk.StringVar()
        self.__collator = tk.StringVar()
        self.__reportTime = tk.StringVar()

        self.__testResult1 = tk.BooleanVar()
        self.__testResult2 = tk.BooleanVar()
        self.__testResult3 = tk.BooleanVar()
        self.__testResult4 = tk.BooleanVar()
        self.__testResult5 = tk.BooleanVar()
        self.__testResult6 = tk.BooleanVar()
        self.__testResult7 = tk.BooleanVar()
        self.__testResult8 = tk.BooleanVar()
        self.__totalResult = tk.StringVar()

        self.__entryLabel = [['编号*','文件*'],
                             ['姓名*','年龄*','联系电话*'],
                             ['科室','申请医生','住院号'],
                             ['床号','接收时间','处理时间'],
                             ['地址'],
                             [],
                             ['样本类型', '样本大小', '检测项目'],
                             ['样本质量', '检测者', '复核者']]

        self.__entryLabel1 = [['未见微生物感染','滴虫感染'],
                              ['念珠菌感染','加德纳菌感染'],
                              ['滴虫感染 + 念珠菌感染','滴虫感染 + 加德纳菌感染'],
                              ['加德纳菌感染 + 念珠菌感染','滴虫感染 + 念珠菌感染 + 加德纳菌感染']]

        #数据, 宽度, colunmspan
        self.__entryConfigList = [[self.__serialVar, 10, 1, self.__fileVar, 34, 4],
                                  [self.__nameVar, 10, 1, self.__ageVar, 13, 1, self.__numberVar, 13, 1],
                                  [self.__departmentVar, 10, 1, self.__doctorVar, 13, 1, self.__hospitalNumVar, 13, 1],
                                  [self.__bedNumVar, 10, 1, self.__receiveVar, 13, 1, self.__handleVar, 13, 1],
                                  [self.__addressVar, 33, 3],
                                  [],
                                  [self.__sampleType, 10, 1, self.__sampleSize, 13, 1, self.__testItem, 13, 1],
                                  [self.__sampleQuality, 10, 1, self.__tester, 13, 1, self.__collator, 13, 1]
                                  ]

        self.__resultChkboxList = [[self.__testResult1, self.__testResult2],
                                   [self.__testResult3, self.__testResult4],
                                   [self.__testResult5, self.__testResult6],
                                   [self.__testResult7, self.__testResult8]]

        self.__testResultList = [self.__testResult1, self.__testResult2,
                            self.__testResult3, self.__testResult4,
                            self.__testResult5, self.__testResult6,
                            self.__testResult7, self.__testResult8]

        self.__hospital = tk.StringVar()
        self.hospitalComb = ttk.Combobox(self.rightFrame, textvariable=self.__hospital, values=self.__hospitalList)
        self.hospitalComb.grid(row=0, column=0, columnspan=3, sticky='w', pady=5)
        self.hospitalComb.bind("<<ComboboxSelected>>", self.__hospitalChanged)
        self.hospitalComb.set(self.__curHospital)

        self.searchEntry = tk.Entry(self.rightFrame, width=28)
        self.searchEntry.grid(row=0, column=3, columnspan=3, sticky=tk.W)

        self.__varList = [
            self.__serialVar, self.__nameVar, self.__ageVar, self.__fileVar,
            self.__numberVar, self.__departmentVar, self.__doctorVar,
            self.__hospitalNumVar, self.__bedNumVar, self.__receiveVar,
            self.__handleVar, self.__addressVar, self.__sampleType,
            self.__sampleSize, self.__testItem, self.__sampleQuality,
            self.__tester, self.__collator, self.__totalResult,
            self.__reportTime, self.__hospital
        ]

        out = 0
        for list in self.__entryLabel1:
            inOut = 0
            for label in list:
                tk.Checkbutton(self.rightFrame,
                               text=label,
                               variable=self.__resultChkboxList[out][inOut]).grid(row=out+9,
                                                                                  column=3*inOut,
                                                                                  columnspan=3,
                                                                                  sticky='W')
                inOut += 1
            out += 1

        out = 0
        for list in self.__entryLabel:
            inOut = 0
            for label in list:
                tk.Label(
                    self.rightFrame, text=label).grid(
                    row=out+1, column=2*inOut, sticky=tk.W, pady=5)
                tk.Entry(
                    self.rightFrame,
                    width=self.__entryConfigList[out][3 * inOut + 1],
                    textvariable=self.__entryConfigList[out][3 * inOut]).grid(
                        row=out + 1,
                        column=2 * inOut + 1,
                        columnspan=self.__entryConfigList[out][3 * inOut + 2],
                        sticky=tk.W,
                        pady=5)
                inOut += 1
            out += 1


        '''
        for attri in self.__infoAttri:
            index = self.__infoAttri.index(attri)
            tk.Label(
                self.rightFrame, text=attri).grid(
                    row=index+1, column=0, sticky=tk.E, pady=5)
        self.idEntry = tk.Entry(
            self.rightFrame,
            textvariable=self.idVar,
            validate='key',
            validatecommand=(validateID, '%P'))
        # 输入框以及绑定事件
        self.idEntry.grid(row=1, column=1, sticky=tk.E)
        self.fileEntry = tk.Entry(self.rightFrame, textvariable=self.fileVar)
        self.fileEntry.grid(row=2, column=1, sticky=tk.E)
        self.nameEntry = tk.Entry(self.rightFrame, textvariable=self.nameVar)
        self.nameEntry.grid(row=3, column=1, sticky=tk.E)
        self.ageEntry = tk.Entry(
            self.rightFrame,
            textvariable=self.ageVar,
            validate='key',
            validatecommand=(validateAge, '%P'))
        self.ageEntry.grid(row=4, column=1, sticky=tk.E)
        '''

    def __setupButtons(self):
        tk.Button(
            self.rightFrame, text='搜索', command=self.__search).grid(
                row=0, column=5, pady=5, ipadx=5, sticky=tk.E)
        tk.Button(
            self.rightFrame, text='更新选中', command=self.__updateConfigBtn).grid(
                row=14, column=0, columnspan=2, pady=5, ipadx=5, sticky=tk.W)
        tk.Button(
            self.rightFrame, text='删除选中', command=self.__delConfigBtn).grid(
                row=14, column=3, columnspan=2, pady=5, ipadx=5, sticky=tk.W)
        tk.Button(
            self.rightFrame, text='增加信息', command=self.__genConfigBtn).grid(
                row=14, column=5, columnspan=2, pady=5, padx=10, ipadx=5, sticky=tk.W)
        tk.Button(
            self.rightFrame, text='生成文档', command=self.__genFile).grid(
                row=15, column=0, columnspan=2, pady=5, ipadx=5, sticky=tk.W)
        tk.Button(
            self.rightFrame, text='关闭', command=self.__closeApp).grid(
                row=15, column=3, pady=5, columnspan=2, ipadx=5, sticky=tk.W)

    def updateTable(self, list):
        # 删除原节点
        items = self.table.get_children()
        [self.table.delete(item) for item in items]
        # for _ in map(self.table.delete, self.table.get_children("")):
        #     pass
        index = 0
        for info in list:
            vals = [index+1]
            vals.extend(info[1:self.tableItem])
            self.table.insert("", "end", values=vals)
            index += 1
        self.brush_treeview(self.table)

    def brush_treeview(self, tv):
        """
        改变treeview样式
        :param tv:
        :return:
        """
        if not isinstance(tv, ttk.Treeview):
            raise Exception(
                "argument tv of method bursh_treeview must be instance of ttk.TreeView"
            )
        #=============设置样式=====
        items = tv.get_children()
        for i in range(len(items)):
            if i % 2 == 1:
                tv.item(items[i], tags=('oddrow'))
        tv.tag_configure('oddrow', background='#eeeeff')

    def on_tree_select(self, event):
        print("selected items:")
        for item in self.table.selection():
            item_text = self.table.item(item, "values")
            print(item_text)

    def get_row_value(self, event):
        """
        获取ttk treeview某一个单元格的值（在鼠标事件中）
        :param event:
        :param tree:
        :param col_widths:
        :return:
        """
        if not isinstance(event, tk.Event):
            raise Exception("event must type of Tkinter.Event")
        x = event.x
        y = event.y
        row = self.table.identify_row(y)
        # colunm = tree.identify_column(x)
        vals = self.table.item(row, "values")
        # print(vals)
        if len(vals) > 2:
            self.currentItem = row
            ret = self.readConfig(self.__hospital.get(), vals[1])
            self.__setValListValue(ret)
        else:
            self.currentItem = None

    def update_row_value(self):
        """
        更新某一行的内容
        """
        if self.currentItem != None:
            self.table.set(self.currentItem, column=1, value=self.__serialVar.get())
            self.table.set(
                self.currentItem, column=2, value=self.__nameVar.get())
            self.table.set(self.currentItem, column=3, value=self.__ageVar.get())
        pass

    def delete_row(self):
        """
        删除某行
        """
        if self.currentItem != None:
            self.table.delete(self.currentItem)
            self.currentItem = None
            self.selectID = 0
            self.selectSerial = 0
            for var in self.__varList:
                var.set("")

    def message(self, info: str):
        messagebox.showinfo('提示', info)

    # 验证id输入
    def __serialValidate(self, content):
        if content.isdigit() or content == "":
            return True
        else:
            self.message('请输入数字')
            return False

    def __ageValidate(self, content):
        if content.isdigit() or content == "":
            return True
        else:
            self.message('请输入数字')
            return False

    def __setValListValue(self, queryVal):
        self.selectID = queryVal[0]
        self.selectSerial = queryVal[1]
        print(self.selectID, self.selectSerial)
        for i in range(len(self.__varList)):
            self.__varList[i].set(queryVal[i+1])
        chkBoxVal = self.__totalResult.get().split(',')
        self.__testResult1.set(bool(int(chkBoxVal[0])))
        self.__testResult2.set(bool(int(chkBoxVal[1])))
        self.__testResult3.set(bool(int(chkBoxVal[2])))
        self.__testResult4.set(bool(int(chkBoxVal[3])))
        self.__testResult5.set(bool(int(chkBoxVal[4])))
        self.__testResult6.set(bool(int(chkBoxVal[5])))
        self.__testResult7.set(bool(int(chkBoxVal[6])))
        self.__testResult8.set(bool(int(chkBoxVal[7])))

    def __closeWindow(self):
        if messagebox.askokcancel("关闭", "确认退出？"):
            self.window.destroy()

    ########### 界面 ##############

    ########### 配置文件 ##############
    def __loadAppConfig(self):
        file = open(self.appConfigFile, mode='r')
        file.close()
        self.appConfig.read(self.appConfigFile)
        if not self.appConfig.has_section('current'):
            self.appConfig.add_section('current')
            self.appConfig.set('current', 'hospital', '绵阳市人民医院病理科')

        if self.appConfig.has_section('hospital'):
            self.__hospitalList = [self.appConfig.get('hospital', option) for option in self.appConfig.options('hospital')]
            self.__curHospital = self.appConfig.get('current', 'hospital')
        else:
            self.appConfig.add_section('hospital')
            self.appConfig.set('hospital', '0', '绵阳市人民医院病理科')
            self.__curHospital = '绵阳市人民医院病理科'
        self.__saveAppConfig()

    def __saveAppConfig(self):
        with open(self.appConfigFile, mode='w') as file:
            self.appConfig.write(file)

    def __checkHospital(self, hospital):
        '''
        检查下拉列表是否存在该医院名称，不存在则增加该医院到下拉列表
        :param hospital:
        :return: bool 存在true
        '''
        ret = False
        self.appConfig.read(self.appConfigFile)
        list = [self.appConfig.get('hospital', option) for option in self.appConfig.options('hospital')]
        if hospital in list:
            ret = True
        else:
            maxIndex = int(self.appConfig.options('hospital')[-1])
            print(maxIndex)
            self.appConfig.set('hospital',str(maxIndex+1),hospital)
            self.__hospitalList.append(hospital)
            self.__curHospital = maxIndex+1
            self.__saveAppConfig()
        return ret

    def __genTestResult(self):
        list = [str(int(item.get())) for item in self.__testResultList]
        ret = ','.join(list)
        return ret

    def readConfig(self, hospital, serial):
        sql = 'SELECT * FROM infoBase WHERE serial=? AND hospital=?'
        self.cursor.execute(sql,(serial, hospital))
        vals = self.cursor.fetchall()
        return vals[0]


    def updateConfig(self):
        valStr = []
        # print(len(self.setstr), len(self.__varList))
        for i in range(len(self.setstr)):
            valStr.append(self.setstr[i]+"='"+self.__varList[i].get()+"'")
        # valStr = 'serial=?,name=?,age=?,file=?,number=?,department=?,doctor=?,hospital=?,bednum=?,receive=?,handle=?,address=?,sampleType=?,sampleSize=?,testItem=?,sampleQuality=?,tester=?,collator=?,testResult=?,reportTime=?'
        sql = "UPDATE infoBase SET {0} WHERE id={1}".format(','.join(valStr) ,self.selectID)
        # print(sql)
        self.cursor.execute(sql)
        self.conn.commit()

    def delConfig(self, id):
        sql = 'DELETE FROM infoBase WHERE id=?'
        self.cursor.execute(sql, (id, ))
        self.conn.commit()

    def genConfig(self):
        vals = [item.get() for item in self.__varList]
        sql = 'INSERT INTO infoBase ({0}) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'.format(','.join(self.setstr))
        # print(sql)
        self.cursor.execute(sql, vals)
        self.conn.commit()

    def checkSameSerial(self, hospital, serial):
        '''
        检查是否存在相同的序列号
        :param hospital:
        :param serial:
        :return: 相同则序列号，不相同返回None
        '''
        ret = None
        sql = 'SELECT * FROM infoBase WHERE serial=? and hospital=?'
        self.cursor.execute(sql,(serial, hospital))
        vals = self.cursor.fetchall()
        if len(vals) != 0:
            ret = serial
        return ret

    def getHospitalCase(self, hospital):
        sql = "SELECT * FROM infoBase WHERE hospital='{0}'".format(hospital)
        self.cursor.execute(sql)
        vals = self.cursor.fetchall()
        return vals

    def testConfigGen(self, num):
        for i in range(num):
            for var in self.__varList:
                var.set(str(i))
            self.__varList[-2].set(time.strftime("%Y-%m-%d", time.localtime()))
            self.__varList[-3].set('1,1,1,0,1,1,0,0')
            self.__hospital.set('眉山市人民医院')
            self.genConfig()

    ########### 配置文件 ##############

    ########### 按键操作 ##############
    def __updateConfigBtn(self):
        if self.currentItem is not None:
            ret = self.checkSameSerial(self.__hospital.get(), self.__serialVar.get())
            if self.__checkEmpty() or (ret is not None and ret != self.selectSerial):
                self.message("存在空数据或者当前医院存在重复编码，请修改后再重试。")
            else:
                if not self.__checkHospital(self.hospitalComb.get()):
                    self.hospitalComb["values"] = self.__hospitalList
                self.updateConfig()
                self.update_row_value()
        else:
            self.message('请选择数据。')
            # self.table.insert('', 'end', values=[self.__serialVar, file, name, age])
        pass

    def __delConfigBtn(self):
        if self.currentItem is not None:
            self.delConfig(self.selectID)
            self.updateTable(self.getHospitalCase(self.__hospital.get()))
        else:
            self.message('请选择数据。')

    def __genConfigBtn(self):
        if self.__checkEmpty() or (self.checkSameSerial(self.__hospital.get(), self.__serialVar.get()) is not None):
            self.message("存在空数据或者编码重复，请修改后再操作。")
        else:
            if not self.__checkHospital(self.hospitalComb.get()):
                self.hospitalComb["values"] = self.__hospitalList
            self.__totalResult.set(self.__genTestResult())
            self.genConfig()
            self.updateTable(self.getHospitalCase(self.__hospital.get()))
            self.table.see(self.table.get_children()[-1])
        pass

    def __genFile(self):
        if self.currentItem == None:
            self.message('请选择一个案例')
            return
        self.generateDocx()
        pass

    def __search(self):
        content = self.searchEntry.get()
        if content != "":
            sql = "SELECT * FROM infoBase WHERE serial LIKE '%{}%' or name Like '%{}%' OR number LIKE '%{}%' AND hospital='{}'".format(content, content, content, self.__hospital.get())
            self.cursor.execute(sql)
            vals = self.cursor.fetchall()
            self.updateTable(vals)
        else:
            self.__initTable()

    def __closeApp(self):
        self.__closeWindow()

    def __hospitalChanged(self, *args):
        infoList = self.getHospitalCase(self.__hospital.get())
        self.updateTable(infoList)
        self.appConfig.read(self.appConfigFile)
        self.appConfig.set('current', 'hospital', self.__hospital.get())
        self.__saveAppConfig()

    def __checkEmpty(self):
        ret = False
        for val in self.__varList[:5]:
            if len(val.get()) < 1:
                ret = True
                break
        if int(self.__serialVar.get()) < 0:
            ret = True
        return ret

    ########### 按键操作 ##############

    def generateDocx(self):
        fileName = '病例文件/{}.docx'.format(self.__fileVar.get())
        # with MailMerge(r'template/template.docx') as document:
        #     for p in document.get_merge_fields():
        #         if p =='file':
        #             document.merge({p:caseInfo.file})
        #             print(caseInfo.file)
        #     document.merge({'file': caseInfo.file})
        #     document.write(fileName)
        doc = DocxTemplate(r'template/case_template.docx')
        # for p in doc.paragraphs:
        #     for index in range(len(self.setstr)):
        #         old_text = self.setstr[index]
        #         if old_text in p.text:
        #             print(p.text)
        #             inline = p.runs
        #             for i in inline:
        #                 print(i.text)
        #                 if old_text in i.text:
        #                     text = i.text.replace(old_text, self.__varList[index].get())
        #                     i.text = text
        #                     print(p.text, i.text, old_text, self.__varList[index].get())
        # doc.save(fileName)
        # result = self.__testResultList
        checked = ['template/unchecked.png', 'template/checked.png']
        result = [checked[int(item.get())] for item in self.__testResultList]
        context = {
            'name': self.__nameVar.get(),
            'serial': self.__serialVar.get(),
            'age': self.__ageVar.get(),
            'file': self.__fileVar.get(),
            'number': self.__numberVar.get(),
            'department': self.__departmentVar.get(),
            'doctor': self.__doctorVar.get(),
            'hospitalnum': self.__hospitalNumVar.get(),
            'bednum': self.__bedNumVar.get(),
            'receive': self.__receiveVar.get(),
            'handle': self.__handleVar.get(),
            'address': self.__addressVar.get(),
            'sampletype': self.__sampleType.get(),
            'samplesize': self.__sampleSize.get(),
            'testitem': self.__testItem.get(),
            'samplequality': self.__sampleQuality.get(),
            'tester': self.__tester.get(),
            'collator': self.__collator.get(),
            'reporttime': self.__reportTime.get(),
            'hospital': self.__hospital.get(),
            'r1': InlineImage(doc, result[0], height=Mm(4)),
            'r2': InlineImage(doc, result[1], height=Mm(4)),
            'r3': InlineImage(doc, result[2], height=Mm(4)),
            'r4': InlineImage(doc, result[3], height=Mm(4)),
            'r5': InlineImage(doc, result[4], height=Mm(4)),
            'r6': InlineImage(doc, result[5], height=Mm(4)),
            'r7': InlineImage(doc, result[6], height=Mm(4)),
            'r8': InlineImage(doc, result[7], height=Mm(4))
        }
        doc.render(context)
        # doc.replace_pic('unchecked.png', 'template/checked.png')
        doc.save(fileName)


if __name__ == '__main__':
    try:
        if not os.path.exists('病例文件'):
            os.makedirs('病例文件')
        if not os.path.exists('config/settings'):
            os.makedirs('config/settings')
        if not os.path.exists('config/database'):
            os.makedirs('config/database')
        if not os.path.exists('config/log'):
            os.makedirs('config/log')
        app = CaseApp()

        # app.testConfigGen(1000)
        #app.getCase(0, 20)

    except configparser.DuplicateSectionError as e:
        error = open('config/log/log.txt', 'w+')
        error.write(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())+'\n'+e.message)
        error.close()
        print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())+'\n'+e.message)
