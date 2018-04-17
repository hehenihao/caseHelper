import base64
import configparser
import os
import shutil
import sqlite3
import time
import tkinter as tk
from tkinter import messagebox, ttk

from mailmerge import MailMerge

from icon import img

os.chdir(os.path.abspath(os.path.dirname(__file__)))
# print(os.getcwd())
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
    def __init__(self, file: str):
        self.config = configparser.ConfigParser()
        self.appConfig = configparser.ConfigParser()
        self.file = file
        self.tableItem = 4
        self.__infoAttri = ("编号", "文件", "姓名", "年龄","电话",
                            "申请医生", "住院号", "床位号", "接收时间",
                            "处理时间", "地址") # 表格数据
        self.__infoAttriEn = ('id', "file", 'name', 'age', 'number', 'department',
                              'doctor', 'hospitalnum', 'bednum', 'receive',
                              'handle', 'address')
        self.__anchor = ('center', 'w', 'w', 'center')
        self.__loadAppConfig()
        self.__setupWidget()

    ########### 界面 ##############
    def __setupWidget(self):
        self.window = tk.Tk()
        self.window.title('Case Helper')
        icon = open('app.ico', 'wb+')
        icon.write(base64.b64decode(img))
        icon.close()
        self.window.iconbitmap("app.ico")
        os.remove("app.ico")
        # self.window.geometry('600x400')
        # w, h = self.window.maxsize()
        w, h = 890, 522
        # 获取屏幕 宽、高
        ws = self.window.winfo_screenwidth()
        hs = self.window.winfo_screenheight()
        # 计算 x, y 位置
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2) - 20
        self.window.geometry('{}x{}+{}+{}'.format(w, h, int(x), int(y)))
        # self.window.resizable(0, 0)  # 防止窗口调整大小

        listNameLbl = tk.Label(
            self.window, text='数据列表', font=('微软雅黑', 10), width=15)
        listNameLbl.grid(row=0, column=0)
        tk.Label(self.window, text='基本信息', font=('微软雅黑', 10), width=15).grid(row=0, column=1)

        self.leftFrame = tk.Frame(self.window, width=420, height=h - 30)
        self.rightFrame = tk.Frame(
            self.window, width=470, height=h - 30)  #, bg='red'
        self.leftFrame.grid(row=1, column=0, padx=5)
        self.rightFrame.grid(row=1, column=1, sticky=tk.N)

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
        treeItemWidth = int(420 / self.tableItem)
        index = 0
        for attri in self.__infoAttri[:self.tableItem]:
            self.table.column(
                attri, width=treeItemWidth,
                anchor=self.__anchor[index])  #表示列,不显示
            self.table.heading(attri, text=attri)  #显示表头
            index += 1
        self.table.grid(row=0, column=0, sticky=tk.NS, pady=5)
        self.vbar.grid(row=0, column=1, sticky=tk.NS)
        self.__initTable()
        self.currentItem = None
        self.__setupInputs()
        self.__setupButtons()
        self.window.mainloop()

    def __initTable(self):
        self.config.read(self.file)
        infoList = list({})
        for section in self.config.sections():
            info = self.readConfig(section)
            infoList.append(info)
        self.updateTable(infoList)

    def __setupInputs(self):
        self.__idVar = tk.StringVar()
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
        self.__entryLabel = [['编号','文件'],
                             ['姓名','年龄','联系电话'],
                             ['科室','申请医生','住院号'],
                             ['床号','接收时间','处理时间'],
                             ['地址']]
        #数据, 宽度, colunmspan
        self.__entryConfigList = [[self.__idVar, 10, 1, self.__fileVar, 34, 4],
                            [self.__nameVar, 10, 1, self.__ageVar, 13, 1, self.__numberVar, 13, 1],
                            [self.__departmentVar, 10, 1, self.__doctorVar, 13, 1, self.__hospitalNumVar, 13, 1],
                            [self.__bedNumVar, 10, 1, self.__receiveVar, 13, 1, self.__handleVar, 13, 1],
                            [self.__addressVar, 33, 3]]

        self.__varList = [
            self.__idVar, self.__fileVar, self.__nameVar, self.__ageVar,
            self.__numberVar, self.__departmentVar, self.__doctorVar,
            self.__hospitalNumVar, self.__bedNumVar, self.__receiveVar,
            self.__handleVar, self.__addressVar
        ]
        '''
        self.__entryList = [[self.__idEntry, self.__fileEntry],
                            [self.__nameEntry,self.__ageEntry, self.__numberEntry],
                            [self.__departmentEntry, self.__doctorEntry, self.__bedNumEntry, self.__hospitalNumEntry],
                            [self.__receiveEntry, self.__handleEntry],
                            [self.__addressEntry]]
                            
        validateID = self.window.register(self.__idValidate)
        validateAge = self.window.register(self.__ageValidate)
        '''

        self.__hospital = tk.StringVar()
        self.hospitalComb = ttk.Combobox(self.rightFrame, textvariable=self.__hospital, values=self.__hospitalList)
        self.hospitalComb.grid(row=0, column=0, columnspan=3, sticky='w', pady=5)
        self.hospitalComb.bind("<<ComboboxSelected>>", self.__hospitalChanged)
        self.hospitalComb.current(self.__curHospital)

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
        self.searchEntry = tk.Entry(self.rightFrame, width=28)
        self.searchEntry.grid(row=0, column=3, columnspan=3, sticky=tk.W)
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
            self.rightFrame, text='更新选中', command=self.__updateConfigBtn).grid(
                row=6, column=0, columnspan=2, pady=5, ipadx=5, sticky=tk.W)
        tk.Button(
            self.rightFrame, text='删除选中', command=self.__delConfigBtn).grid(
                row=6, column=3, columnspan=2, pady=5, ipadx=5, sticky=tk.W)
        tk.Button(
            self.rightFrame, text='增加信息', command=self.__genConfigBtn).grid(
                row=6, column=5, columnspan=2, pady=5, padx=10, ipadx=5, sticky=tk.W)
        tk.Button(
            self.rightFrame, text='生成文档', command=self.__genFile).grid(
                row=7, column=0, columnspan=2, pady=5, ipadx=5, sticky=tk.W)
        tk.Button(
            self.rightFrame, text='搜索', command=self.__search).grid(
                row=0, column=5, pady=5, ipadx=5, sticky=tk.E)
        tk.Button(
            self.rightFrame, text='关闭', command=self.__closeApp).grid(
                row=7, column=3, pady=5, columnspan=2,ipadx=5, sticky=tk.W)

    def updateTable(self, list):
        # 删除原节点
        items = self.table.get_children()
        [self.table.delete(item) for item in items]
        # for _ in map(self.table.delete, self.table.get_children("")):
        #     pass
        for info in list:
            self.table.insert("", "end", values=info[0:self.tableItem])
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
            id = vals[0]
            print(id)
            self.currentItem = row
        else:
            self.currentItem = None

    def update_row_value(self):
        """
        更新某一行的内容
        """
        if self.currentItem != None:
            self.table.set(self.currentItem, column=0, value=self.__idVar.get())
            self.table.set(
                self.currentItem, column=1, value=self.__fileVar.get())
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
            for var in self.__varList:
                    var.set("")

    def message(self, info: str):
        messagebox.showinfo('提示', info)

    # 验证id输入
    def __idValidate(self, content):
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

    ########### 界面 ##############

    ########### 配置文件 ##############
    def __loadAppConfig(self):
        self.appConfig.read('app.txt')
        self.__hospitalList = [self.appConfig.get('hospital', option) for option in self.appConfig.options('hospital')]
        self.__curHospital = int(self.appConfig.get('current', 'hospital'))

    def readConfig(self, index):
        id = str(index)
        self.config.read(self.file, encoding='gbk')
        vals = list({})
        if self.config.has_section(id):
            vals.append(id)
            for attri in self.__infoAttriEn[1:]:
                # name = self.config.get(id, 'name')
                # age = self.config.get(id, 'age')
                vals.append(self.config.get(id, attri))
        return vals

    def updateConfig(self):
        self.config.read(self.file, encoding='gbk')
        if not self.config.has_section(self.__idVar.get()):
            self.config.add_section(self.__idVar.get())
        # index = 1
        # for attri in self.__infoAttriEn[1:]:
        #     self.config.set(caseInfo.id, attri, caseInfo.vals[index])
        #     index += 1
        self.config.set(self.__idVar.get(), 'name', self.__nameVar.get())
        self.config.set(self.__idVar.get(), 'age', self.__ageVar.get())
        self.config.set(self.__idVar.get(), 'file', self.__fileVar.get())
        self.config.set(self.__idVar.get(), 'number', self.__numberVar.get())
        self.config.set(self.__idVar.get(), 'department', self.__departmentVar.get())
        self.config.set(self.__idVar.get(), 'doctor', self.__doctorVar.get())
        self.config.set(self.__idVar.get(), 'hospitalNum', self.__hospitalNumVar.get())
        self.config.set(self.__idVar.get(), 'bedNum', self.__bedNumVar.get())
        self.config.set(self.__idVar.get(), 'receive', self.__receiveVar.get())
        self.config.set(self.__idVar.get(), 'handle', self.__handleVar.get())
        self.config.set(self.__idVar.get(), 'address', self.__addressVar.get())
        with open(self.file, 'w') as configFile:
            self.config.write(configFile)

    def delConfig(self, id):
        self.config.read(self.file, encoding='gbk')
        if self.config.has_section(str(id)):
            self.config.remove_section(str(id))
            with open(self.file, 'w') as configFile:
                self.config.write(configFile)

    def genConfig(self):
        ret = False
        self.config.read(self.file, encoding='gbk')
        if self.config.has_section(self.__idVar.get()):
            self.message('id已经存在')
        else:
            self.updateConfig()
            ret = True
        return ret

    def sortConfig(self):
        self.config.read(self.file, encoding='gbk')
        for item in self.config.items():
            for k in item:
                print(k)

    def testConfigGen(self, num):
        for i in range(num):
            for var in self.__varList:
                var.set(str(i))
            self.updateConfig()

    ########### 配置文件 ##############

    ########### 按键操作 ##############
    def __updateConfigBtn(self):
        id, file, name, age = self.__idVar.get(), self.__fileVar.get(), self.__nameVar.get(), self.__ageVar.get()
        if len(str(id)) > 0 and len(name) > 0 and len(str(age)) > 0:
            self.updateConfig()
            self.update_row_value()
        else:
            self.message('数据不完整')

    def __delConfigBtn(self):
        print(self.__idVar.get())
        self.delConfig(self.__idVar.get())
        self.delete_row()

    def __genConfigBtn(self):
        id, file, name, age = self.__idVar.get(), self.__fileVar.get(
        ), self.__nameVar.get(), self.__ageVar.get()
        if len(str(id)) > 0 and len(name) > 0 and len(str(age)) > 0:
            if self.genConfig():
                self.table.insert('', 'end', values=[id, file, name, age])
            self.table.see(self.table.get_children()[-1])
        pass

    def __genFile(self):
        if self.currentItem == None:
            self.message('请选择一个案例')
            return
        vals = self.table.item(self.currentItem, "values")
        info = CaseInfo(vals)
        self.generateDocx(info)
        pass

    def __search(self):
        content = self.searchEntry.get()
        if content != "":
            infoList = list({})
            for item in self.table.get_children():
                vals = self.table.item(item, 'values')
                if vals[1].find(content) > -1 or vals[2].find(content) > -1:
                    info = CaseInfo(list(vals))
                    infoList.append(info)
            self.updateTable(infoList)

        else:
            self.__initTable()

    def __closeApp(self):
        self.window.quit()

    def __hospitalChanged(self, *args):
        print(self.__hospital.get(), args)

    ########### 按键操作 ##############

    def generateDocx(self, caseInfo: CaseInfo):
        fileName = '病例文件/{}_{}_{}.docx'.format(caseInfo.id, caseInfo.name,
                                               caseInfo.file)
        with MailMerge(r'template/template.docx') as document:
            # for p in document.get_merge_fields():
            #     if p =='file':
            #         document.merge({p:caseInfo.file})
            #         print(caseInfo.file)
            document.merge({'file': caseInfo.file})
            document.write(fileName)


if __name__ == '__main__':
    try:
        if not os.path.exists('病例文件'):
            os.makedirs('病例文件')
        app = CaseApp('case.txt')
        # app.testConfigGen(1000)
    except configparser.DuplicateSectionError as e:
        error = open('config/log/log.txt', 'w+')
        error.write(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())+'\n'+e.message)
        error.close()
        print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())+'\n'+e.message)
