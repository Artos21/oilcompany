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

class function(object):
    def showMessage(self, mes):
        self.msgBox.setText(mes)
        self.msgBox.exec()

    def checkInput1(self):
        try:
            self.cursor = self.connection.cursor()
            self.cursor.execute("SELECT `mining_station`.`id` FROM `mining_station`;")
            self.results = self.cursor.fetchall()
            self.number_station = int(self.line1.text())
            self.rows = len(self.results)
            for i in range(self.rows):
                if self.number_station == self.results[i][0]:
                    self.bull = True
            if self.bull == False:
                raise Exception
            self.bull = False
        except Exception:
            self.showMessage("Введите номер станции корректно")
            return False
        return True

    def checkInput2(self):
        try:
            self.number_station = int(self.line2.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите номер станции корректно")
            return False

        return True

    def checkInput3(self):
        try:
            self.number_station = int(self.lineEdit.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер хранилища")
            return False

        try:
            self.number_station = float(self.lineEdit_2.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный объем нефти")
            return False

        try:
            self.number_station = int(self.lineEdit_46.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер типа приемки")
            return False

        try:
            self.number_station = float(self.lineEdit_47.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер статуса")
            return False
        try:
            self.number_station = float(self.lineEdit_3.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер станции")
            return False
        try:
            self.number_station = float(self.lineEdit_4.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер работника")
            return False
        return True

    def checkInput4(self):
        try:
            self.number_station = int(self.lineEdit_13.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер распоряжения")
            return False
        try:
            self.number_station = int(self.lineEdit_9.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер хранилища")
            return False

        try:
            self.number_station = float(self.lineEdit_10.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный объем нефти")
            return False

        try:
            self.number_station = int(self.lineEdit_11.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер типа приемки")
            return False

        try:
            self.number_station = float(self.lineEdit_12.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер статуса")
            return False
        try:
            self.number_station = float(self.lineEdit_48.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер станции")
            return False
        try:
            self.number_station = float(self.lineEdit_49.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер работника")
            return False
        return True

    def checkInput5(self):
        try:
            self.number_station = int(self.lineEdit_50.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите номер распоряжения корректно")
            return False

        return True

    def checkInput6(self):
        try:
            self.number_station = int(self.lineEdit_14.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер хранилища")
            return False

        try:
            self.number_station = float(self.lineEdit_15.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный объем нефти")
            return False

        try:
            self.number_station = int(self.lineEdit_51.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер типа отгрузки")
            return False

        try:
            self.number_station = float(self.lineEdit_52.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер статуса")
            return False
        try:
            self.number_station = float(self.lineEdit_16.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный код заказчика")
            return False
        try:
            self.number_station = float(self.lineEdit_17.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер работника")
            return False
        return True

    def checkInput7(self):
        try:
            self.number_station = int(self.lineEdit_22.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер распоряжения")
            return False
        try:
            self.number_station = int(self.lineEdit_18.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер хранилища")
            return False

        try:
            self.number_station = float(self.lineEdit_19.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный объем нефти")
            return False

        try:
            self.number_station = int(self.lineEdit_20.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер типа отгрузки")
            return False

        try:
            self.number_station = float(self.lineEdit_21.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер статуса")
            return False
        try:
            self.number_station = float(self.lineEdit_53.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный код заказчика")
            return False
        try:
            self.number_station = float(self.lineEdit_54.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите корректный номер работника")
            return False
        return True

    def checkInput8(self):
        try:
            self.number_station = int(self.lineEdit_55.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите номер распоряжения корректно")
            return False

        return True

    def checkInput9(self):
        try:
            self.number_station = int(self.line1_3.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите номер хранилища корректно")
            return False

        return True

    def checkInput10(self):
        try:
            self.number_station = int(self.line2_4.text())
            if self.number_station <= 0:
                raise Exception
        except Exception:
            self.showMessage("Введите номер хранилища корректно")
            return False

        return True

    def checkDate(self):

        self.date_1 = self.dateEdit.dateTime().toString('yyyy-MM-dd')
        self.date_2 = self.dateEdit_2.dateTime().toString('yyyy-MM-dd')
        print(self.date_1, self.date_2)

    def checkconnection(self):
        try:
            self.cursor = self.connection.cursor()
            self.cursor.execute("SELECT VERSION()")
            self.results = self.cursor.fetchone()
            if self.results:
                return True
            else:
                return False
        except Exception:
            self.showMessage("Нет соединения")
            return False
        return True




    def recordDB_1(self):
        if not self.checkconnection():
            return
        self.cursor = self.connection.cursor()
        self.cursor.execute("SELECT `mining_station`.`id`, `mining_station`.`number_station`, `field`.`name`, `status`.`name`, `mode`.`name`, `fuel_type`.`name`, `oil_prod`.`quantity` FROM `mining_station` LEFT JOIN `field` ON `mining_station`.`id_field` = `field`.`id` LEFT JOIN `status` ON `mining_station`.`id_status` = `status`.`id` LEFT JOIN `mode` ON `mining_station`.`id_mode` = `mode`.`id` LEFT JOIN `fuel_type` ON `mining_station`.`fuel_type_id` = `fuel_type`.`id` LEFT JOIN `oil_prod` ON `oil_prod`.`id_station` = `mining_station`.`id` GROUP BY `mining_station`.`id`;")
        self.results = self.cursor.fetchall()

        if len(self.results) == 0:
             self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            self.tableWidget.setRowCount(self.rows)
            self.tableWidget.setColumnCount(self.columns)
            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget.setItem(i, j, item)


    def recordDB_2(self):
        self.cursor = self.connection.cursor()
        self.cursor.execute("SELECT `oil_prod`.`id`, `oil_prod`.`date`, `oil_prod`.`quantity`, `mining_station`.`number_station`, `field`.`name` FROM `oil_prod` LEFT JOIN `mining_station` ON `oil_prod`.`id_station` = `mining_station`.`id` LEFT JOIN `field` ON `mining_station`.`id_field` = `field`.`id`;")
        self.results = self.cursor.fetchall()
        # print(self.results)
        if len(self.results) == 0:
             self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            self.tableWidget_2.setRowCount(self.rows)
            self.tableWidget_2.setColumnCount(self.columns)
            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget_2.setItem(i, j, item)

    def recordDB_3(self):
        self.cursor = self.connection.cursor()
        self.cursor.execute("SELECT `supply`.`id`, `supply`.`quantity`, `supply`.`date`, `type_supply`.`type`, `state`.`state`, `tank_farm`.`number_tank`, `workers`.`tab_number`, `mining_station`.`id` FROM `supply` LEFT JOIN `type_supply` ON `supply`.`type_supply` = `type_supply`.`id` LEFT JOIN `state` ON `supply`.`state_id` = `state`.`id` LEFT JOIN `tank_farm` ON `supply`.`tank_id` = `tank_farm`.`id` LEFT JOIN `workers` ON `supply`.`workers_id` = `workers`.`id` LEFT JOIN `mining_station` ON `supply`.`mining_station_id` = `mining_station`.`id`;")
        self.results = self.cursor.fetchall()
        # print(self.results)
        if len(self.results) == 0:
             self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            self.tableWidget_3.setRowCount(self.rows)
            self.tableWidget_3.setColumnCount(self.columns)
            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget_3.setItem(i, j, item)

    def recordDB_4(self):
        self.cursor = self.connection.cursor()
        self.cursor.execute("SELECT `entrepot`.`id`, `entrepot`.`date`, `entrepot`.`quantity`, `tank_farm`.`id` FROM `entrepot` LEFT JOIN `tank_farm` ON `entrepot`.`tank_id` = `tank_farm`.`id`;")
        self.results = self.cursor.fetchall()
        # print(self.results)
        if len(self.results) == 0:
             self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            self.tableWidget_4.setRowCount(self.rows)
            self.tableWidget_4.setColumnCount(self.columns)
            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget_4.setItem(i, j, item)

    def recordDB_5(self):
        self.cursor = self.connection.cursor()
        self.cursor.execute("SELECT `delivery`.`id`, `delivery`.`quantity`, `delivery`.`date`, `type_delivery`.`type`, `state`.`state`, `tank_farm`.`id`, `workers`.`id`, `customer`.`id` FROM `delivery` LEFT JOIN `type_delivery` ON `delivery`.`type_delivery_id` = `type_delivery`.`id` LEFT JOIN `state` ON `delivery`.`state_id` = `state`.`id` LEFT JOIN `tank_farm` ON `delivery`.`tank_id` = `tank_farm`.`id` LEFT JOIN `workers` ON `delivery`.`workers_id` = `workers`.`id` LEFT JOIN `customer` ON `delivery`.`customer_id` = `customer`.`id`;")
        self.results = self.cursor.fetchall()
        # print(self.results)
        if len(self.results) == 0:
             self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            self.tableWidget_5.setRowCount(self.rows)
            self.tableWidget_5.setColumnCount(self.columns)
            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget_5.setItem(i, j, item)

    def recordDB_6(self):
        self.cursor = self.connection.cursor()
        self.cursor.execute("SELECT `tank_farm`.`id`, `tank_farm`.`location`, `tank_farm`.`quantity_max`, `fuel_type`.`name`FROM `tank_farm` LEFT JOIN `fuel_type` ON `tank_farm`.`fuel_type_id` = `fuel_type`.`id`;")
        self.results = self.cursor.fetchall()
        # print(self.results)
        # print(self.buff)
        if len(self.results) == 0:
             self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            self.tableWidget_16.setRowCount(self.rows)
            self.tableWidget_16.setColumnCount(self.columns)

            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget_16.setItem(i, j, item)

    def pushButton1(self):
        if not self.checkInput1():
            return
        if not self.checkconnection():
            return
        self.table_query = f'UPDATE mining_station SET id_status = 1 WHERE mining_station.id ="{self.number_station}"'
        self.cursor = self.connection.cursor()
        self.cursor.execute(self.table_query)
        self.connection.commit()
        self.recordDB_1()

    def pushButton1_2(self):
        if not self.checkInput1():
            return
        self.buff = self.comboBox.currentText()
        if self.buff == "Активный":
            self.mode = 1
        elif self.buff == "Дежурный":
            self.mode = 2
        elif self.buff == "Аварийный":
            self.mode = 3
        elif self.buff == "Производительный":
            self.mode = 4
        elif self.buff == "Неактивный":
            self.mode = 5
        self.table_query = f'UPDATE mining_station SET id_mode = "{self.mode}" WHERE mining_station.id ="{self.number_station}"'
        self.cursor = self.connection.cursor()
        self.cursor.execute(self.table_query)
        self.connection.commit()
        self.recordDB_1()

    def pushButton1_3(self):
        if not self.checkInput1():
            return
        self.cursor = self.connection.cursor()
        self.cursor.execute(F'''SELECT `mining_station`.`id`, `mining_station`.`number_station`, `field`.`name`, `status`.`name`, `mode`.`name`, `fuel_type`.`name` FROM `mining_station` LEFT JOIN `field` ON `mining_station`.`id_field` = `field`.`id` LEFT JOIN `status` ON `mining_station`.`id_status` = `status`.`id` LEFT JOIN `mode` ON `mining_station`.`id_mode` = `mode`.`id` LEFT JOIN `fuel_type` ON `mining_station`.`fuel_type_id` = `fuel_type`.`id` WHERE `mining_station`.`id` = "{self.number_station}";''')
        self.results = self.cursor.fetchall()
        if len(self.results) == 0:
            self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget.setItem(i, j, item)
            self.tableWidget.setRowCount(self.rows)
            self.tableWidget.setColumnCount(self.columns)
            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget.setItem(i, j, item)

    def pushButton2(self):
        if not self.checkInput1():
            return

        self.table_query = f'UPDATE mining_station SET id_status = 2 WHERE mining_station.id ="{self.number_station}"'
        self.cursor = self.connection.cursor()
        self.cursor.execute(self.table_query)
        self.connection.commit()
        self.recordDB_1()

    def pushButton2_1(self):
        if not self.checkInput9():
            return

        self.cursor = self.connection.cursor()
        self.cursor.execute(F'''SELECT `tank_farm`.`id`, `tank_farm`.`location`, `tank_farm`.`quantity_max`, `fuel_type`.`name`FROM `tank_farm` LEFT JOIN `fuel_type` ON `tank_farm`.`fuel_type_id` = `fuel_type`.`id` WHERE `tank_farm`.`id` ="{self.number_station}";''')
        self.results = self.cursor.fetchall()
        if len(self.results) == 0:
            self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            self.tableWidget_16.setRowCount(self.rows)
            self.tableWidget_16.setColumnCount(self.columns)

            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget_16.setItem(i, j, item)

    def pushButton2_2(self):
        self.recordDB_6()

    def pushButton3_1(self):
        if not self.checkInput3():
            return

        self.table_query = f'''INSERT INTO `supply` (`id`, `quantity`, `date`, `type_supply`, `state_id`, `tank_id`, `workers_id`, `mining_station_id`) VALUES (NULL, "{float(self.lineEdit_2.text())}", "{self.dateEdit_3.dateTime().toString('yyyy-MM-dd')}", "{int(self.lineEdit_46.text())}", "{int(self.lineEdit_47.text())}", "{int(self.lineEdit.text())}", "{int(self.lineEdit_4.text())}", "{int(self.lineEdit_3.text())}")'''
        self.cursor = self.connection.cursor()
        self.cursor.execute(self.table_query)
        self.connection.commit()
        self.recordDB_3()

    def pushButton3_2(self):
        if not self.checkInput4():
            return

        self.table_query = f'''UPDATE `supply` SET  `quantity` = "{float(self.lineEdit_10.text())}", `date` = "{self.dateEdit_5.dateTime().toString('yyyy-MM-dd')}", `type_supply` = "{int(self.lineEdit_11.text())}", `state_id` = "{int(self.lineEdit_12.text())}", `tank_id` = "{int(self.lineEdit_9.text())}", `workers_id` = "{int(self.lineEdit_49.text())}", `mining_station_id` = "{int(self.lineEdit_48.text())}" WHERE `supply`.`id` = {int(self.lineEdit_13.text())}'''
        self.cursor = self.connection.cursor()
        self.cursor.execute(self.table_query)
        self.connection.commit()
        self.recordDB_3()

    def pushButton3_3(self):
        if not self.checkInput5():
            return

        self.table_query = f'''DELETE FROM `supply` WHERE `supply`.`id` = {int(self.lineEdit_50.text())}'''
        self.cursor = self.connection.cursor()
        self.cursor.execute(self.table_query)
        self.connection.commit()
        self.recordDB_3()
    def pushButton3_4(self):
        if not self.checkInput10():
            return

        self.cursor = self.connection.cursor()
        self.cursor.execute(f'''SELECT `entrepot`.`id`, `entrepot`.`date`, `entrepot`.`quantity`, `tank_farm`.`id` FROM `entrepot` LEFT JOIN `tank_farm` ON `entrepot`.`tank_id` = `tank_farm`.`id` WHERE `tank_farm`.`id` = "{self.number_station}" ORDER BY date DESC LIMIT 1 ''')
        self.results = self.cursor.fetchall()
        print(self.results)
        if len(self.results) == 0:
            self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            self.tableWidget_4.setRowCount(self.rows)
            self.tableWidget_4.setColumnCount(self.columns)
            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget_4.setItem(i, j, item)

    def pushButton3_5(self):

        self.recordDB_4()

    def pushButton3_6(self):
        if not self.checkInput10():
            return
        # self.checkDate()
        self.cursor = self.connection.cursor()
        self.cursor.execute(f'''SELECT `entrepot`.`id`, `entrepot`.`date`, `entrepot`.`quantity`, `tank_farm`.`id` FROM `entrepot` LEFT JOIN `tank_farm` ON `entrepot`.`tank_id` = `tank_farm`.`id` WHERE `tank_farm`.`id` = "{self.number_station}" AND `entrepot`.`date` BETWEEN "{self.dateEdit_7.dateTime().toString('yyyy-MM-dd')}" AND "{self.dateEdit_8.dateTime().toString('yyyy-MM-dd')}"''')
        self.results = self.cursor.fetchall()
        if len(self.results) == 0:
            self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            self.tableWidget_4.setRowCount(self.rows)
            self.tableWidget_4.setColumnCount(self.columns)
            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget_4.setItem(i, j, item)

    def pushButton4_1(self):
        if not self.checkInput6():
            return

        self.table_query = f'''INSERT INTO `delivery` (`id`, `date`, `quantity`, `tank_id`, `state_id`, `workers_id`, `customer_id`, `type_delivery_id`) VALUES (NULL, "{self.dateEdit_6.dateTime().toString('yyyy-MM-dd')}", "{float(self.lineEdit_15.text())}", "{int(self.lineEdit_14.text())}", "{int(self.lineEdit_52.text())}", "{int(self.lineEdit_17.text())}", "{int(self.lineEdit_16.text())}", "{int(self.lineEdit_51.text())}")'''
        self.cursor = self.connection.cursor()
        self.cursor.execute(self.table_query)
        self.connection.commit()
        self.recordDB_5()

    def pushButton4_2(self):
        if not self.checkInput7():
            return

        self.table_query = f'''UPDATE `delivery` SET `date` = {self.dateEdit_9.dateTime().toString('yyyy-MM-dd')}, `quantity` = "{float(self.lineEdit_19.text())}", `tank_id` = "{int(self.lineEdit_18.text())}", `state_id` = "{int(self.lineEdit_21.text())}", `workers_id` = "{int(self.lineEdit_54.text())}", `customer_id` = "{int(self.lineEdit_53.text())}", `type_delivery_id` = "{int(self.lineEdit_20.text())}" WHERE `delivery`.`id` = "{int(self.lineEdit_22.text())}"'''
        self.cursor = self.connection.cursor()
        self.cursor.execute(self.table_query)
        self.connection.commit()
        self.recordDB_5()

    def pushButton4_3(self):
        if not self.checkInput8():
            return

        self.table_query = f'''DELETE FROM `delivery` WHERE `delivery`.`id` = {int(self.lineEdit_55.text())}'''
        self.cursor = self.connection.cursor()
        self.cursor.execute(self.table_query)
        self.connection.commit()
        self.recordDB_5()


    def pushButton5(self):
        if not self.checkInput2():
            return

        self.cursor = self.connection.cursor()
        self.cursor.execute(f'SELECT `oil_prod`.`id`, `oil_prod`.`date`, `oil_prod`.`quantity`, `mining_station`.`id`, `field`.`name` FROM `oil_prod` LEFT JOIN `mining_station` ON `oil_prod`.`id_station` = `mining_station`.`id` LEFT JOIN `field` ON `mining_station`.`id_field` = `field`.`id` WHERE `mining_station`.`id` = "{self.number_station}" ORDER BY date DESC LIMIT 1 ;')
        self.results = self.cursor.fetchall()
        # print(self.results)
        if len(self.results) == 0:
             self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            self.tableWidget_2.setRowCount(self.rows)
            self.tableWidget_2.setColumnCount(self.columns)
            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget_2.setItem(i, j, item)

    def pushButton6(self):
        self.recordDB_2()

    def pushButton6_2(self):
        if not self.checkInput2():
            return
        self.checkDate()
        self.cursor = self.connection.cursor()
        self.cursor.execute(f'''SELECT `oil_prod`.`id`, `oil_prod`.`date`, `oil_prod`.`quantity`, `mining_station`.`id`, `field`.`name` FROM `oil_prod` LEFT JOIN `mining_station` ON `oil_prod`.`id_station` = `mining_station`.`id` LEFT JOIN `field` ON `mining_station`.`id_field` = `field`.`id` WHERE `mining_station`.`id` = "{self.number_station}" AND `oil_prod`.`date` BETWEEN "{self.date_1}" AND "{self.date_2}";''')
        self.results = self.cursor.fetchall()

        if len(self.results) == 0:
             self.showMessage("Нет данных")
        else:
            self.rows = len(self.results)
            self.columns = len(self.results[0])
            self.tableWidget_2.setRowCount(self.rows)
            self.tableWidget_2.setColumnCount(self.columns)
            for i in range(self.rows):
                for j in range(self.columns):
                    item = QTableWidgetItem("{}".format(self.results[i][j]))
                    self.tableWidget_2.setItem(i, j, item)

    def pushButton7(self):
        self.cursor = self.connection.cursor()
        self.cursor.execute(
            "SELECT `mining_station`.`id`, `mining_station`.`number_station`, `field`.`name`, `status`.`name`, `mode`.`name`, `fuel_type`.`name` FROM `mining_station` LEFT JOIN `field` ON `mining_station`.`id_field` = `field`.`id` LEFT JOIN `status` ON `mining_station`.`id_status` = `status`.`id` LEFT JOIN `mode` ON `mining_station`.`id_mode` = `mode`.`id` LEFT JOIN `fuel_type` ON `mining_station`.`fuel_type_id` = `fuel_type`.`id`;")
        self.results = self.cursor.fetchall()
        self.rows = len(self.results)
        self.columns = len(self.results[0])
        print(self.results)
        print(self.rows)
        print(self.columns)
        book = openpyxl.Workbook()
        sheet = book.active
        row = 2
        col = 0
        sheet['A1'] = 'ID'
        sheet['B1'] = 'Номер станции'
        sheet['C1'] = 'Месторождение'
        sheet['D1'] = 'Статус'
        sheet['E1'] = 'Ружим работы'
        sheet['F1'] = 'Тип топлива'
        sheet[row][0].value = self.results[1][1]

        # for i in range(self.rows):
        #     sheet[row][0].value = self.results[i][0]
        #     sheet[row][1].value = self.results[i][1]
        #     sheet[row][2].value = self.results[i][2]
        #     sheet[row][3].value = self.results[i][3]
        #     sheet[row][4].value = self.results[i][4]
        #     sheet[row][5].value = self.results[i][5]
        #     row+=1
        try:

            for i in range(self.rows):
                for k in range(self.columns):
                    sheet[row][k].value = self.results[i][k]
                row+=1
            self.showMessage("Успешно")
        except:
            print(Exception)

        book.save("doo.xlsx")
        book.close()