from PyQt5 import QtWidgets, uic, QtCore, QtGui
from pyqtgraph import PlotWidget, plot
from PyQt5.QtWidgets import QVBoxLayout,QMessageBox
import pandas as pd
import sys
import os
import serial
import threading
import res
import openpyxl

class LoginWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = uic.loadUi("login.ui", self)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.ui.comboBox.setEditable(True)
        self.ui.comboBox.setCurrentIndex(-1)
        self.ui.comboBox.setCurrentText("Choose Account Type")
        self.ui.comboBox.addItem("Patient")
        self.ui.comboBox.addItem("Nurse")
        self.ui.comboBox.addItem("Admin")
        self.ui.stackedWidget.setCurrentIndex(0)
        self.mode = 0
        self.isLogin = True

        self.ui.pushButton_12.clicked.connect(self.minimize)
        self.ui.pushButton_10.clicked.connect(self.minimize)
        self.ui.pushButton_9.clicked.connect(self.exit)
        self.ui.pushButton_14.clicked.connect(self.exit)
        self.ui.pushButton_2.clicked.connect(self.SwitchLogin)
        self.ui.pushButton_6.clicked.connect(self.SwitchLogin)
        self.ui.pushButton.clicked.connect(self.Login)
        self.ui.pushButton_5.clicked.connect(self.Register)
        self.ui.comboBox.currentIndexChanged.connect(self.UpdateCombo)
        self.ui.comboBox.currentTextChanged.connect(self.comboBoxFixer)

    def UpdateCombo(self):
        self.ui.comboBox.setEditable(False)
        win = MainWindow()
        win.show()

    def SwitchLogin(self):
        self.isLogin = not self.isLogin
        if self.isLogin:
            self.ui.stackedWidget.setCurrentIndex(0)
        else:
            self.ui.stackedWidget.setCurrentIndex(1)
            print(self.ui.comboBox.currentText())

    def Login(self):
        username = self.ui.lineEdit.text()
        password = self.ui.lineEdit_2.text()
        users = pd.read_excel("assets\\Database\\users.xlsx", sheet_name=['admins','patients','nurses'])
        admins = users['admins']
        patients = users['patients']
        nurses = users['nurses']
        found = False
        real_pw = ''
        c=0
        for user in [admins,patients,nurses]:
            if found:
                break
            c += 1
            for i,p in zip(user['username'],user['password']):
                if username == i:
                    found = True
                    real_pw = p
                    break

        if found and (real_pw == password):
            self.ui.label_17.setText("Welcome!")
            self.mode = c
        elif found and (real_pw != password):
            self.ui.label_17.setText("Wrong Password, Please Try Again.")
        elif not found:
            self.ui.label_17.setText("Username doesn't exist.")

    def comboBoxFixer(self):
        if self.ui.comboBox.isEditable() == True and self.ui.comboBox.currentText != "Choose Account Type":
            self.ui.comboBox.setCurrentText("Choose Account Type")


    def Register(self):
        username = self.ui.lineEdit_5.text()
        password = self.ui.lineEdit_6.text()
        confirm = self.ui.lineEdit_9.text()
        users = pd.read_excel("assets\\Database\\users.xlsx", sheet_name=['admins','patients','nurses'])
        admins = users['admins']
        patients = users['patients']
        nurses = users['nurses']
        requests = pd.read_excel("assets\\Database\\requests.xlsx",sheet_name=['requests'])
        chosen = self.ui.comboBox.currentText()
        found = False
        for user in [admins,patients,nurses,requests['requests']]:
            if found:
                break
            for i in user['username']:
                if username == i:
                    found = True
                    break
        if ' ' in username:
            self.ui.label_18.setText("Username can't contain spaces.")
        elif username == '' or password == '' or confirm == '' or chosen == "Choose Account Type":
            self.ui.label_18.setText("Please fill all requirements.")
        elif password != confirm:
            self.ui.label_18.setText("Passwords don't match.")
        elif found:
            self.ui.label_18.setText("Username is taken.")
        else:
            book = openpyxl.load_workbook('assets\\Database\\requests.xlsx')
            reqs = book.active
            reqs.append([username,password,chosen])
            book.save(os.getcwd() + "\\assets\\Database\\requests.xlsx")
            self.ui.label_18.setText("Registered, waiting for admin confirmation.")
        
    def exit(self):
        self.close()

    def minimize(self):
        self.showNormal()
        self.showMinimized()


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = uic.loadUi("GUI.ui", self)