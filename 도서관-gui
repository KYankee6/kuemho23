-- coding: utf-8 --
Form implementation generated from reading ui file 'mywindow.ui'
Created by: PyQt5 UI code generator 5.9.2
WARNING! All changes made in this file will be lost!
'''
Author=양준혁
Author : 양준혁
Author-email : didzl1231@naver.com
Version : 1.00
Build Tools : Python 3.7
Date it was made: 2020-02-04
last update : 2020-02-04
'''
import win32com.client
import os
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
bookList = {}
excel = win32com.client.Dispatch("Excel.Application")

def getExcel(path):
if not os.path.isfile(path):
exit()
wb=excel.Workbooks.Open(path)
ws=wb.ActiveSheet
for i in range(10000):
if ws.Cells(4+i, 2).Value == None : break
bookList[ws.Cells(4+i,2).Value]=4+i
excel.Quit()

class Ui_MainWindow(object):
path = r"C:\Users\82105\Desktop\Library.xlsx"
i = 0
def setupUi(self, MainWindow):
# 엑셀 오픈
if not os.path.isfile(path):
exit()
self.wb = excel.Workbooks.Open(path)
self.ws = self.wb.ActiveSheet
excel.Visible = True

    #트리거



    # 메인 윈도우 창
    MainWindow.setObjectName("MainWindow")
    MainWindow.resize(800, 600)
    self.centralwidget = QtWidgets.QWidget(MainWindow)
    self.centralwidget.setObjectName("centralwidget")

    # 라인 에디터 (입력창)
    self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
    self.lineEdit.setGeometry(QtCore.QRect(0, 50, 581, 81))
    self.lineEdit.setObjectName("lineEdit")

    # label 시리얼 번호 입력
    self.label = QtWidgets.QLabel(self.centralwidget)
    self.label.setGeometry(QtCore.QRect(10, 20, 100, 15))
    self.label.setObjectName("label")

    # label  집계현황
    self.label_2 = QtWidgets.QLabel(self.centralwidget)
    self.label_2.setGeometry(QtCore.QRect(43, 224, 721, 151))
    self.label_2.setObjectName("label_2")

    # label 카운터 라벨
    self.label_3 = QtWidgets.QLabel(self.centralwidget)
    self.label_3.setGeometry(QtCore.QRect(0, 180, 64, 15))
    self.label_3.setObjectName("label_3")

    # 카운터
    self.label_4 = QtWidgets.QLabel(self.centralwidget)
    self.label_4.setGeometry(QtCore.QRect(590, 442, 181, 81))
    self.label_4.setObjectName("label_4")

    #선들
    self.line = QtWidgets.QFrame(self.centralwidget)
    self.line.setGeometry(QtCore.QRect(0, 420, 811, 16))
    self.line.setFrameShape(QtWidgets.QFrame.HLine)
    self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
    self.line.setObjectName("line")
    self.line_2 = QtWidgets.QFrame(self.centralwidget)
    self.line_2.setGeometry(QtCore.QRect(0, 160, 811, 16))
    self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
    self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
    self.line_2.setObjectName("line_2")


    #버튼 - 입력 (시리얼 입력이후)
    self.pushButton = QtWidgets.QPushButton(self.centralwidget)
    self.pushButton.setGeometry(QtCore.QRect(600, 50, 181, 81))
    self.pushButton.setObjectName("pushButton")

    # 엑셀화 버튼
    self.pushButton_1 = QtWidgets.QPushButton(self.centralwidget)
    self.pushButton_1.setGeometry(QtCore.QRect(100, 442, 181, 81))
    self.pushButton_1.setObjectName("pushButton_1")

    # 메뉴바 상단 (미사용)
    MainWindow.setCentralWidget(self.centralwidget)
    self.menubar = QtWidgets.QMenuBar(MainWindow)
    self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 26))
    self.menubar.setObjectName("menubar")
    MainWindow.setMenuBar(self.menubar)

    # 상태바 하단(미사용)
    self.statusbar = QtWidgets.QStatusBar(MainWindow)
    self.statusbar.setObjectName("statusbar")
    MainWindow.setStatusBar(self.statusbar)

    self.retranslateUi(MainWindow)
    QtCore.QMetaObject.connectSlotsByName(MainWindow)

    _translate = QtCore.QCoreApplication.translate
    # 입력 버튼 이벤트 처리
    self.pushButton.setText(_translate("MainWindow", "입력"))
    self.pushButton.clicked.connect(self.lineEdit_returned)

    # 엑셀화 버튼 이벤트 처리
    self.pushButton_1.setText(_translate("MainWindow", "엑셀화/꼭 눌러주세요"))
    self.pushButton_1.clicked.connect(self.btn1_clicked)

    # 입력창
    self.lineEdit.setText(_translate("MainWindow", ""))
    self.lineEdit.returnPressed.connect(self.lineEdit_returned)

def retranslateUi(self, MainWindow):
    _translate = QtCore.QCoreApplication.translate
    MainWindow.setWindowTitle(_translate("MainWindow", "소장자료확인기"))
    self.label.setText(_translate("MainWindow", "시리얼 입력"))
    self.label_3.setText(_translate("MainWindow", "집계현황"))

def lineEdit_returned(self):

    self.statusbar.showMessage(self.lineEdit.text())
    if self.lineEdit.text() in bookList:
        self.i+=1
        self.ws.Cells(bookList[self.lineEdit.text()], 20).Value = "O"
        self.lineEdit.clear()
        self.label_2.setText("존재합니다")
        self.label_2.setStyleSheet("color:#0000FF;font-size:25px")
        self.label_4.setText(str(self.i))
        self.label_4.setStyleSheet("font-size:25px")
    else :
        self.label_2.setText("존재하지않습니다")
        self.label_2.setStyleSheet("color:#FF0000;font-size:25px")
        self.lineEdit.clear()

def btn1_clicked(self):
    excel.Quit()
    sys.exit(app.exec_())
if name == "main":
path=r"C:\Users\82105\Desktop\Library.xlsx"
getExcel(path)
app = QtWidgets.QApplication(sys.argv)
MainWindow = QtWidgets.QMainWindow()
ui = Ui_MainWindow()
ui.setupUi(MainWindow)
MainWindow.show()
app.exec_()
