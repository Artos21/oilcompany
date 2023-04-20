import function
from Design_py import window
from PyQt5 import (QtWidgets, QtCore)
from PyQt5.QtWidgets import QTableWidgetItem
import pymysql.cursors
import time
import pymysql
import mysql.connector
from getpass import getpass
from mysql.connector import connect, Error
import pandas as pd
import os
import openpyxl



class MainWindow(QtWidgets.QWidget, window.Ui_Form, function.function):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.msgBox = QtWidgets.QMessageBox()
        self.setFixedSize(self.width(), self.height())
        self.number_station = 0
        self.date_1 = ''
        self.date_2 = ''
        self.bull = False
        self.mode = 0
        self.cursor = 0
        self.results = 0
        self.buff = ""
        self.new_arr = []
        self.rows = 0
        self.columns = 0
        self.table_query = ""

        try:
            self.connection = pymysql.connect(host='127.0.0.1',port=3306,user='root',password='root',database='oil_company',charset='utf8')
            self.recordDB_1()
            self.recordDB_2()
            self.recordDB_3()
            self.recordDB_4()
            self.recordDB_5()
            self.recordDB_6()
            self.button1.clicked.connect(self.pushButton1)
            self.button2.clicked.connect(self.pushButton2)
            self.button3.clicked.connect(self.pushButton1_3)
            self.button4.clicked.connect(self.pushButton1_2)
            self.button5.clicked.connect(self.pushButton5)
            self.button6.clicked.connect(self.pushButton6)
            self.button4_2.clicked.connect(self.pushButton7)
            self.button4_7.clicked.connect(self.recordDB_1)
            self.button6_2.clicked.connect(self.pushButton6_2)
            self.button6_5.clicked.connect(self.pushButton3_1)
            self.button6_6.clicked.connect(self.pushButton3_2)
            self.button6_11.clicked.connect(self.pushButton3_3)
            self.button6_12.clicked.connect(self.pushButton4_1)
            self.button6_13.clicked.connect(self.pushButton4_2)
            self.button6_14.clicked.connect(self.pushButton4_3)
            self.button4_9.clicked.connect(self.pushButton2_2)
            self.button3_3.clicked.connect(self.pushButton2_1)
            self.button5_4.clicked.connect(self.pushButton3_4)
            self.button6_8.clicked.connect(self.pushButton3_5)
            self.button6_7.clicked.connect(self.pushButton3_6)
        except Exception:
            self.showMessage("Нет соединения с базой данных")
