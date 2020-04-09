#!Anaconda/anaconda/python
#coding: utf-8

# pyinstaller -F -w -i .\pump.ico .\main_search.py --hidden-import PyQt5.sip

import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from searchv1 import Ui_search
import xlrd


class MyMainForm(QMainWindow, Ui_search):
    def __init__(self, parent=None):
        super(MyMainForm, self).__init__(parent)
        self.setupUi(self)
        # 添加 检索 按钮的信号和槽
        self.search_button.clicked.connect(self.searchStart)
        # 添加 清空 按钮信号和槽
        self.clear_button.clicked.connect(self.clearResult)
        # 读取总数据库
        self.myDataBase = xlrd.open_workbook('./data/database.xls')


    # 根据用户选择加载数据库中对应的sheet
    def loadDatabase(self):
        # 确定要检索的型号，这里model的值对应于xls中sheet的索引
        if self.SB_button.isChecked():
            self.model = 0
        elif self.SBL_button.isChecked():
            self.model = 1
        elif self.WQ_button.isChecked():
            self.model = 2
        elif self.xiaofang_button.isChecked():
            self.model = 3
        else:
            self.model = 4
        self.mySheetBase = self.myDataBase.sheet_by_index(self.model)
        # 获取所有参数
        self.allModels = self.mySheetBase.col_values(0, start_rowx=1)
        self.allFlows = self.mySheetBase.col_values(1, start_rowx=1)
        self.allLifts = self.mySheetBase.col_values(2, start_rowx=1)


    # 开始检索
    def searchStart(self):
        # 重新根据型号加载数据，防止第二次检索时改变型号
        self.loadDatabase()
        # 获取用户输入的数据
        flow = self.flow_input.text()
        lift = self.lift_input.text()
        if flow.replace('.','',1).isdigit()==False or lift.replace('.','',1).isdigit()==False:
            QMessageBox.information(self, "格式错误", "请输入数字！")
            return 0

        if self.precise_button.isChecked():
            self.precesePattern(float(flow), float(lift))
        else:
            self.fuzzyPattern(float(flow), float(lift))

    # 精确检索方法
    def precesePattern(self, flows, lifts):
        flowAddressIndex = []
        matchIndex = []
        self.clearResult()
        # 查找流量和扬程都匹配的型号索引
        for i in range(len(self.allFlows)):
            if self.allFlows[i] == flows:
                flowAddressIndex.append(i)
        for j in range(len(flowAddressIndex)):
            if self.allLifts[flowAddressIndex[j]] == lifts:
                matchIndex.append(flowAddressIndex[j])
        if len(matchIndex) == 0:
            QMessageBox.information(self, "抱歉", "没有查找到对应型号！")
            return 0
        else:
            self.printResult(matchIndex)


    # 模糊检索方法
    def fuzzyPattern(self, flows, lifts):
        flowAddressIndex = []
        matchIndex = []
        self.clearResult()
        # 设置模糊范围
        flowWide = self.setWide(flows, 15)
        liftWide = self.setWide(lifts, 10)
        # 查找流量和扬程都模糊匹配的型号索引
        for i in range(len(self.allFlows)):
            if flowWide['down'] <= self.allFlows[i] <= flowWide['up']:
                flowAddressIndex.append(i)
        for j in range(len(flowAddressIndex)):
            if liftWide['down'] <= self.allLifts[flowAddressIndex[j]] <= liftWide['up']:
                matchIndex.append(flowAddressIndex[j])
        if len(matchIndex) == 0:
            QMessageBox.information(self, "抱歉", "没有查找到对应型号！")
            return 0
        else:
            self.printResult(matchIndex)


    # 在信息框中输出检索到的型号及其信息
    def printResult(self, index):
        title = ['型号', '流量', '扬程', '转速', '功率']
        titlePrint = ["{:^12}".format(z) for z in title]
        self.result_text.append(''.join(titlePrint))
        self.result_text.append("-" * 70)
        #self.result_text.append("型号\t\t|\t流量\t|\t扬程\t|\t转速\t|\t功率\t\n")
        for i in index:
            myRow = self.mySheetBase.row_values(i+1)
            myStr = [str(k) for k in myRow]
            myStr2 = ["{:^14}".format(p) for p in myStr]
            self.result_text.append(''.join(myStr2))
            self.result_text.append("-"*70)


    # 把基数按照浮动范围设置为字典类型的数据
    def setWide(self, baseVal, wide):
        upVal = baseVal + wide
        downVal = baseVal - wide
        if downVal < 0:
            downVal = 0
        tempVal = {'down': downVal, 'up': upVal}
        return tempVal

    # 清除信息框中的数据，防止与下一次数据重复
    def clearResult(self):
        self.result_text.clear()



if __name__ == "__main__":
    app = QApplication(sys.argv)
    # init the frame
    myWin = MyMainForm()
    # show the frame in windows
    myWin.show()
    sys.exit(app.exec_())
