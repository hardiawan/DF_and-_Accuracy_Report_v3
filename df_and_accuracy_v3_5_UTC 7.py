from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QMessageBox
from datetime import datetime, timedelta

from pymavlink import mavutil
import os
import pandas as pd
import geopandas as gpd
import pyproj
from shapely.geometry import Point
import PyPDF2
import shutil
import csv
from shapely.wkt import loads
import matplotlib.pyplot as plt
import simplekml
import seaborn as sns
from fpdf import FPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import io
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, NamedStyle
from openpyxl.styles import Alignment, Font
from openpyxl.utils.cell import get_column_letter
from openpyxl.drawing.image import Image
from openpyxl import load_workbook

import time
from html2image import Html2Image
import requests
import folium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os
import time
import matplotlib.dates as mdates
import json

class Ui_MainWindow(object):
    def __init__(self):
        # Initialize a dictionary to store form input history
        self.form_input_history = {}

        # Define a path for saving the history file (JSON)
        self.history_file_path = "form_input_history.json"

        # Load any existing history from the file
        self.load_form_input_history()

    def save_form_input_history(self):
        # Save the current form input history to a JSON file
        with open(self.history_file_path, 'w') as history_file:
            json.dump(self.form_input_history, history_file)

    def load_form_input_history(self):
        # Check if the history file exists, and if so, load it
        if os.path.exists(self.history_file_path):
            with open(self.history_file_path, 'r') as history_file:
                self.form_input_history = json.load(history_file)
                
                
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1370, 750)                                    #Resolusi window sopwer
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 1370, 750))       #Resolusi menubar
        font = QtGui.QFont()
        font.setPointSize(10)
        self.tabWidget.setFont(font)
        self.tabWidget.setObjectName("tabWidget")
        
        self.setupDailyReportTab()
        self.setupMergeReportTab()
        
        self.progressBar = QtWidgets.QProgressBar(MainWindow)
        self.progressBar.setGeometry(QtCore.QRect(750, 640, 600, 21))
        self.progressBar.setMinimum(0)
        self.progressBar.setMaximum(100)
        self.progressBar.setObjectName("progressBar")

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 26))           
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def setupDailyReportTab(self):
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")

        labels = [
            ("label", "Input Location Folder Bin", 20, 10),
            ("label_2", "Input Location Folder Waypoints", 20, 35),
            ("label_3", "Choose Location Folder Output", 20, 60),
            ("label_4", "Choose Company", 20, 85),
            ("label_5", "Choose Estate", 20, 110),
            ("label_6", "Project Location", 20, 135),
            ("label_8", "Drone ID     e.g. D16-01", 20, 160),
            ("label_9", "Choose UTM Zone", 20, 185),
            ("label_10", "Choose Accuration Value", 20, 210),
            
            ("label_10a", "Arrival Time to Location", 700, 10),
            ("label_10b", "Height From Crop (m)", 700, 35),
            ("label_10c", "Flying Speed (m/s)", 700, 60),
            ("label_10d", "Swath (m)", 700, 85),
            ("label_team", "Team Name", 700, 130),
            # ("label_10e", "Drone Type", 700, 170),
            ("label_10f", "Herbicide Ration (%)", 700, 195),
            ("label_10g", "Chemical Type", 700, 220),
            ("label_10h", "Chemical Mixture", 700, 275),
            ("label_10i", "Driver", 700, 420),
            ("label_10j", "Car Plate No", 700, 445),

            ("label_10k", "Herbicide Time Ready", 1040, 10),
            ("label_10l", "PIC", 1040, 35),
            ("label_10m", "Pilot", 1040, 60),
            ("label_10n", "Co-Pilot/Coordinator", 1040, 85),
            ("label_10o", "GCS", 1040, 110),
            ("label_10p", "Support", 1040, 135),
            ("label_10q", "Nozzle Type (Color)", 1040, 170),
            ("label_10r", "District Name", 1040, 195),
            ("label_10s", "District PIC", 1040, 220),
            ("label_10t", "Rotation", 1040, 245),
            ("label_10u", "Plot Number 1", 1040, 300),
            ("label_10v", "Plot Number 2", 1040, 324),
            ("label_10w", "Plot Number 3", 1040, 350),
            ("label_10x", "Plot Number 4", 1040, 375),
            ("label_10y", "Plot Number 5", 1040, 400),
            ("label_Plotnum", "Block    SPH   Dose", 1180, 275),
            ("label_10z1", "Countur Degree", 1040, 445),
            # ("label_10z2", "Date", 1040, 445),
            
            ("label_10z3", "Officer:", 700, 490),
            ("label_10z4", "Mission:", 700, 515),
            ("label_10z5", "Purpose:", 700, 560),

            ("label_10_situation", "Situation", 160, 250),
            ("label_10_solution", "Solution", 480, 250),
            ("label_10_working", "Working Time", 310, 430),

            ("label_11", "*Please fill be carefull*", 1000, 650),
        ]

        for name, text, x, y in labels:
            label = QtWidgets.QLabel(self.tab)
            label.setGeometry(QtCore.QRect(x, y, 280, 20))
            font = QtGui.QFont()
            font.setPointSize(10)
            label.setFont(font)
            label.setObjectName(name)
            label.setText(text)

        self.setupDailyReportTabInputs()
        self.tabWidget.addTab(self.tab, "")

    def setupDailyReportTabInputs(self):
        self.pushButton_4 = QtWidgets.QPushButton(self.tab)
        self.pushButton_4.setGeometry(QtCore.QRect(540, 60, 140, 20))
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_4.setText("Folder Output")
        # self.pushButton_4.clicked.connect(lambda: self.selectFolder("Folder Output", self.lineEdit_4))
        self.pushButton_4.clicked.connect(lambda: self.pushButton_handler4("Folder Output"))

        self.pushButton_3 = QtWidgets.QPushButton(self.tab)
        self.pushButton_3.setGeometry(QtCore.QRect(540, 35, 140, 20))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.setText("Folder .Waypoints")
        # self.pushButton_3.clicked.connect(lambda: self.selectFolder("Folder .Waypoint", self.lineEdit_3))   #Orinya
        self.pushButton_3.clicked.connect(lambda: self.pushButton_handler3("Folder .Waypoints"))

        self.pushButton_2 = QtWidgets.QPushButton(self.tab)
        self.pushButton_2.setGeometry(QtCore.QRect(540, 10, 140, 20))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.setText("Folder .Bin")
        # self.pushButton_2.clicked.connect(lambda: self.selectFolder("Folder .Log", self.lineEdit_2))   #Orinya
        self.pushButton_2.clicked.connect(lambda: self.pushButton_handler2("Folder .Bin"))

        self.pushButton = QtWidgets.QPushButton(self.tab)
        self.pushButton.setGeometry(QtCore.QRect(1180, 645, 80, 25))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.setText("Process")
        self.pushButton.clicked.connect(lambda: self.pushButton_handler1("Process"))

        self.lineEdit_combo6 = QtWidgets.QLineEdit(self.tab)
        self.lineEdit_combo6.setGeometry(QtCore.QRect(540, 135, 140, 20))
        self.lineEdit_combo6.setObjectName("lineEdit_combo6")        

        self.comboBox_2 = self.createComboBox(self.tab, ["SMART Tbk.", "SD GUTHRIE", "Others"], 540, 85)
        self.comboBox_1 = self.createComboBox(self.tab, [], 540, 110)  # Start with an empty comboBox
        # Connect comboBox_2 to the function that updates comboBox_1
        self.comboBox_2.currentIndexChanged.connect(self.updateComboBox2)
        # Initialize comboBox_1 with default selection based on the first item in comboBox_2
        self.updateComboBox1()
        
    def createComboBox(self, parent, items, x, y):
        comboBox = QtWidgets.QComboBox(parent)
        comboBox.setGeometry(QtCore.QRect(x, y, 150, 22))
        comboBox.addItems(items)
        return comboBox
    def updateComboBox1(self):


        # self.comboBox_3 = self.createComboBox(self.tab, ["47N", "47S", "48N", "48S", "49N", "49S", "50S", "51S", "52S"], 540, 185)
        # self.comboBox_3a = self.createComboBox(self.tab, ["< 10cm", "< 15cm", "< 20cm", "< 25cm", "< 30cm", "< 35cm",  "< 40cm"], 540, 210)
        
        if os.path.exists(self.history_file_path):
            with open(self.history_file_path, 'r') as history_file:
                self.form_input_history = json.load(history_file)
                
            self.comboBox_3 = self.createComboBox(self.tab, ["47N", "47S", "48N", "48S", "49N", "49S", "50S", "51S", "52S"], 540, 185)
            index_a = self.comboBox_3.findText(self.form_input_history["comboBox_3"], QtCore.Qt.MatchFixedString)
            if index_a >= 0:
                self.comboBox_3.setCurrentIndex(index_a)
            
            self.comboBox_3a = self.createComboBox(self.tab, ["< 10cm", "< 15cm", "< 20cm", "< 25cm", "< 30cm", "< 35cm",  "< 40cm"], 540, 210)
            index2 = self.comboBox_3a.findText(self.form_input_history["comboBox_3a"], QtCore.Qt.MatchFixedString)
            if index2 >= 0:
                self.comboBox_3a.setCurrentIndex(index2)
            
            self.lineEdit_2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_2.setGeometry(QtCore.QRect(210, 10, 320, 20))
            self.lineEdit_2.setObjectName("lineEdit_2")

            self.lineEdit_3 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_3.setGeometry(QtCore.QRect(210, 35, 320, 20))
            self.lineEdit_3.setObjectName("lineEdit_3")

            self.lineEdit_4 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_4.setGeometry(QtCore.QRect(210, 60, 320, 20))
            self.lineEdit_4.setObjectName("lineEdit_4")

            self.lineEdit_4a = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_4a.setGeometry(QtCore.QRect(540, 160, 140, 20))
            self.lineEdit_4a.setObjectName("lineEdit_4a")
            self.lineEdit_4a.setText(self.form_input_history["lineEdit_4a"])

            self.lineEdit_10a = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10a.setGeometry(QtCore.QRect(850, 10, 140, 20))
            self.lineEdit_10a.setObjectName("lineEdit_10a")
            self.lineEdit_10a.setText(self.form_input_history["lineEdit_10a"])

            self.lineEdit_10b = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10b.setGeometry(QtCore.QRect(850, 35, 140, 20))
            self.lineEdit_10b.setObjectName("lineEdit_10b")
            self.lineEdit_10b.setText(self.form_input_history["lineEdit_10b"])

            self.lineEdit_10c = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10c.setGeometry(QtCore.QRect(850, 60, 140, 20))
            self.lineEdit_10c.setObjectName("lineEdit_10c")
            self.lineEdit_10c.setText(self.form_input_history["lineEdit_10c"])

            self.lineEdit_10d = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10d.setGeometry(QtCore.QRect(850, 85, 140, 20))
            self.lineEdit_10d.setObjectName("lineEdit_10d")
            self.lineEdit_10d.setText(self.form_input_history["lineEdit_10d"])
            
            self.lineEdit_combo6 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_combo6.setGeometry(QtCore.QRect(540, 135, 140, 20))
            self.lineEdit_combo6.setObjectName("lineEdit_combo6")
            self.lineEdit_combo6.setText(self.form_input_history["lineEdit_combo6"])
            
            self.lineEdit_team = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_team.setGeometry(QtCore.QRect(850, 130, 140, 20))
            self.lineEdit_team.setObjectName("lineEdit_team")
            self.lineEdit_team.setText(self.form_input_history["lineEdit_team"])
            
            self.lineEdit_10f = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10f.setGeometry(QtCore.QRect(850, 195, 140, 20))
            self.lineEdit_10f.setObjectName("lineEdit_10f")
            self.lineEdit_10f.setText(self.form_input_history["lineEdit_10f"])

            self.comboBox_10g = self.createComboBox(self.tab, ["Insecticide", "Herbicide", "Fertilizer", "Fungicide", "Larvacide"], 850, 220)
            index = self.comboBox_10g.findText(self.form_input_history["comboBox_10g"], QtCore.Qt.MatchFixedString)
            if index >= 0:
                self.comboBox_10g.setCurrentIndex(index)
                
                
            self.lineEdit_10h = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10h.setGeometry(QtCore.QRect(850, 275, 140, 20))
            self.lineEdit_10h.setObjectName("lineEdit_10h")
            self.lineEdit_10h.setText(self.form_input_history["lineEdit_10h"])

            self.lineEdit_10i = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10i.setGeometry(QtCore.QRect(850, 420, 140, 20))
            self.lineEdit_10i.setObjectName("lineEdit_10i")
            self.lineEdit_10i.setText(self.form_input_history["lineEdit_10i"])

            self.lineEdit_10j = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10j.setGeometry(QtCore.QRect(850, 445, 140, 20))
            self.lineEdit_10j.setObjectName("lineEdit_10j")
            self.lineEdit_10j.setText(self.form_input_history["lineEdit_10j"])

            self.lineEdit_10k = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10k.setGeometry(QtCore.QRect(1180, 10, 140, 20))
            self.lineEdit_10k.setObjectName("lineEdit_10k")
            self.lineEdit_10k.setText(self.form_input_history["lineEdit_10k"])

            self.lineEdit_10l = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10l.setGeometry(QtCore.QRect(1180, 35, 140, 20))
            self.lineEdit_10l.setObjectName("lineEdit_10l")
            self.lineEdit_10l.setText(self.form_input_history["lineEdit_10l"])

            self.lineEdit_10m = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10m.setGeometry(QtCore.QRect(1180, 60, 140, 20))
            self.lineEdit_10m.setObjectName("lineEdit_10m")
            self.lineEdit_10m.setText(self.form_input_history["lineEdit_10m"])

            self.lineEdit_10n = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10n.setGeometry(QtCore.QRect(1180, 85, 140, 20))
            self.lineEdit_10n.setObjectName("lineEdit_10n")
            self.lineEdit_10n.setText(self.form_input_history["lineEdit_10n"])

            self.lineEdit_10o = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10o.setGeometry(QtCore.QRect(1180, 110, 140, 20))
            self.lineEdit_10o.setObjectName("lineEdit_10o")
            self.lineEdit_10o.setText(self.form_input_history["lineEdit_10o"])

            self.lineEdit_10p = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10p.setGeometry(QtCore.QRect(1180, 135, 140, 20))
            self.lineEdit_10p.setObjectName("lineEdit_10p")
            self.lineEdit_10p.setText(self.form_input_history["lineEdit_10p"])

            self.lineEdit_10q = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10q.setGeometry(QtCore.QRect(1180, 170, 140, 20))
            self.lineEdit_10q.setObjectName("lineEdit_10q")
            self.lineEdit_10q.setText("Cone")  # Set default value to "Cone"

            self.lineEdit_10r = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10r.setGeometry(QtCore.QRect(1180, 195, 140, 20))
            self.lineEdit_10r.setObjectName("lineEdit_10r")
            self.lineEdit_10r.setText(self.form_input_history["lineEdit_10r"])

            self.lineEdit_10s = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10s.setGeometry(QtCore.QRect(1180, 220, 140, 20))
            self.lineEdit_10s.setObjectName("lineEdit_10s")
            self.lineEdit_10s.setText(self.form_input_history["lineEdit_10s"])

            # Create and set up QSpinBox
            self.spinBox_10t = QtWidgets.QSpinBox(self.tab)
            self.spinBox_10t.setGeometry(QtCore.QRect(1180, 245, 35, 20))
            self.spinBox_10t.setFont(QtGui.QFont("Arial", 10))

            # Create and set up QComboBox
            self.comboBox_10t = QtWidgets.QComboBox(self.tab)
            self.comboBox_10t.setGeometry(QtCore.QRect(1225, 245, 95, 20))
            self.comboBox_10t.setFont(QtGui.QFont("Arial", 10))
            self.comboBox_10t.addItems(["Rotation", "New", "Spot Reservice"])

            self.lineEdit_10u = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10u.setGeometry(QtCore.QRect(1180, 305, 50, 20))
            self.lineEdit_10u.setObjectName("lineEdit_10u")
            self.lineEdit_10u.setText(self.lineEdit_10u.text().upper())
            self.lineEdit_10u.textChanged.connect(self.on_text_changed5)
            
            self.lineEdit_10u1 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10u1.setGeometry(QtCore.QRect(1235, 305, 40, 20))
            self.lineEdit_10u1.setObjectName("lineEdit_10u1")
            
            self.lineEdit_10u2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10u2.setGeometry(QtCore.QRect(1280, 305, 40, 20))
            self.lineEdit_10u2.setObjectName("lineEdit_10u2")
            
            self.lineEdit_10v = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10v.setGeometry(QtCore.QRect(1180, 330, 50, 20))
            self.lineEdit_10v.setObjectName("lineEdit_10v")
            self.lineEdit_10v.setText(self.lineEdit_10v.text().upper())
            self.lineEdit_10v.setText("**")  
            self.lineEdit_10v.textChanged.connect(self.on_text_changed1)
            
            self.lineEdit_10v1 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10v1.setGeometry(QtCore.QRect(1235, 330, 40, 20))
            self.lineEdit_10v1.setObjectName("lineEdit_10v1")
            self.lineEdit_10v1.setText("**")  
            
            self.lineEdit_10v2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10v2.setGeometry(QtCore.QRect(1280, 330, 40, 20))
            self.lineEdit_10v2.setObjectName("lineEdit_10v2")
            self.lineEdit_10v2.setText("**")  

            self.lineEdit_10w = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10w.setGeometry(QtCore.QRect(1180, 355, 50, 20))
            self.lineEdit_10w.setObjectName("lineEdit_10w")
            self.lineEdit_10w.setText(self.lineEdit_10w.text().upper())
            self.lineEdit_10w.setText("**")
            self.lineEdit_10w.textChanged.connect(self.on_text_changed2)
            
            self.lineEdit_10w1 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10w1.setGeometry(QtCore.QRect(1235, 355, 40, 20))
            self.lineEdit_10w1.setObjectName("lineEdit_10w1")
            self.lineEdit_10w1.setText("**")
            
            self.lineEdit_10w2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10w2.setGeometry(QtCore.QRect(1280, 355, 40, 20))
            self.lineEdit_10w2.setObjectName("lineEdit_10w2")
            self.lineEdit_10w2.setText("**")
            
            self.lineEdit_10x = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10x.setGeometry(QtCore.QRect(1180, 380, 50, 20))
            self.lineEdit_10x.setObjectName("lineEdit_10x")
            self.lineEdit_10x.setText("**")
            self.lineEdit_10x.setText(self.lineEdit_10x.text().upper())
            self.lineEdit_10x.textChanged.connect(self.on_text_changed3)

            self.lineEdit_10x1 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10x1.setGeometry(QtCore.QRect(1235, 380, 40, 20))
            self.lineEdit_10x1.setObjectName("lineEdit_10x1")
            self.lineEdit_10x1.setText("**")
            
            self.lineEdit_10x2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10x2.setGeometry(QtCore.QRect(1280, 380, 40, 20))
            self.lineEdit_10x2.setObjectName("lineEdit_10x2")
            self.lineEdit_10x2.setText("**")


            self.lineEdit_10y = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10y.setGeometry(QtCore.QRect(1180, 405, 50, 20))
            self.lineEdit_10y.setObjectName("lineEdit_10y")
            self.lineEdit_10y.setText("**")
            self.lineEdit_10y.setText(self.lineEdit_10y.text().upper())
            self.lineEdit_10y.textChanged.connect(self.on_text_changed4)
            
            self.lineEdit_10y1 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10y1.setGeometry(QtCore.QRect(1235, 405, 40, 20))
            self.lineEdit_10y1.setObjectName("lineEdit_10y1")
            self.lineEdit_10y1.setText("**")
            
            self.lineEdit_10y2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10y2.setGeometry(QtCore.QRect(1280, 405, 40, 20))
            self.lineEdit_10y2.setObjectName("lineEdit_10y2")
            self.lineEdit_10y2.setText("**")

            self.comboBox_10z1 = self.createComboBox(self.tab, ["10 deg. (Gentle Slope)",  "20 deg.(Contour Slope)"], 1180, 445)
            index1 = self.comboBox_10z1.findText(self.form_input_history["comboBox_10z1"], QtCore.Qt.MatchFixedString)
            if index1 >= 0:
                self.comboBox_10z1.setCurrentIndex(index1)

            self.lineEdit_10z3 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10z3.setGeometry(QtCore.QRect(850, 490, 330, 20))
            self.lineEdit_10z3.setObjectName("lineEdit_10z3")
            self.lineEdit_10z3.setText(self.form_input_history["lineEdit_10z3"])

            self.lineEdit_10z4 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10z4.setGeometry(QtCore.QRect(850, 515, 330, 40))
            self.lineEdit_10z4.setObjectName("lineEdit_10z4")
            self.lineEdit_10z4.setText("Spraying Operation")  # Set default value to "3m"

            self.textEdit_10z5 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_10z5.setGeometry(QtCore.QRect(850, 560, 330, 40))
            self.textEdit_10z5.setObjectName("textEdit_10z5")
            self.textEdit_10z5.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_10z5.setText("To complement the spraying with the ‘palm by palm’ method on plot number ")

            self.textEdit_situation_1 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_situation_1.setGeometry(QtCore.QRect(20, 270, 330, 40))
            self.textEdit_situation_1.setObjectName("textEdit_situation_1")
            self.textEdit_situation_1.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_situation_1.setText(self.form_input_history["textEdit_situation_1"])

            self.textEdit_solution_1 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_solution_1.setGeometry(QtCore.QRect(350, 270, 330, 40))
            self.textEdit_solution_1.setObjectName("textEdit_solution_1")
            self.textEdit_solution_1.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_solution_1.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_solution_1.setText(self.form_input_history["textEdit_solution_1"])

            self.textEdit_situation_2 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_situation_2.setGeometry(QtCore.QRect(20, 310, 330, 40))
            self.textEdit_situation_2.setObjectName("textEdit_situation_2")
            self.textEdit_situation_2.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_situation_2.setText(self.form_input_history["textEdit_situation_2"])

            self.textEdit_solution_2 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_solution_2.setGeometry(QtCore.QRect(350, 310, 330, 40))
            self.textEdit_solution_2.setObjectName("textEdit_solution_2")
            self.textEdit_solution_2.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_solution_2.setText(self.form_input_history["textEdit_solution_2"])

            self.textEdit_situation_3 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_situation_3.setGeometry(QtCore.QRect(20, 350, 330, 40))
            self.textEdit_situation_3.setObjectName("textEdit_situation_3")
            self.textEdit_situation_3.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_situation_3.setText(self.form_input_history["textEdit_situation_3"])

            self.textEdit_solution_3 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_solution_3.setGeometry(QtCore.QRect(350, 350, 330, 40))
            self.textEdit_solution_3.setObjectName("textEdit_solution_3")
            self.textEdit_solution_3.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_solution_3.setText(self.form_input_history["textEdit_situation_3"])

            self.textEdit_situation_4 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_situation_4.setGeometry(QtCore.QRect(20, 390, 330, 40))
            self.textEdit_situation_4.setObjectName("textEdit_situation_4")
            self.textEdit_situation_4.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_situation_4.setText(self.form_input_history["textEdit_situation_4"])

            self.textEdit_solution_4 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_solution_4.setGeometry(QtCore.QRect(350, 390, 330, 40))
            self.textEdit_solution_4.setObjectName("textEdit_solution_4")
            self.textEdit_solution_4.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_solution_4.setText(self.form_input_history["textEdit_solution_4"])

            self.textEdit_worktime_1a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_1a.setGeometry(QtCore.QRect(20, 450, 330, 40))
            self.textEdit_worktime_1a.setObjectName("textEdit_worktime_1a")
            self.textEdit_worktime_1a.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_1a.setText(self.form_input_history["textEdit_worktime_1a"])

            self.textEdit_worktime_1b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_1b.setGeometry(QtCore.QRect(350, 450, 330, 40))
            self.textEdit_worktime_1b.setObjectName("textEdit_worktime_1b")
            self.textEdit_worktime_1b.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_1b.setText(self.form_input_history["textEdit_worktime_1b"])

            self.textEdit_worktime_2a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_2a.setGeometry(QtCore.QRect(20, 480, 330, 40))
            self.textEdit_worktime_2a.setObjectName("textEdit_worktime_2a")
            self.textEdit_worktime_2a.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_2a.setText(self.form_input_history["textEdit_worktime_2a"])

            self.textEdit_worktime_2b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_2b.setGeometry(QtCore.QRect(350, 480, 330, 40))
            self.textEdit_worktime_2b.setObjectName("textEdit_worktime_2b")
            self.textEdit_worktime_2b.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_2b.setText(self.form_input_history["textEdit_worktime_2b"])

            self.textEdit_worktime_3a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_3a.setGeometry(QtCore.QRect(20, 510, 330, 40))
            self.textEdit_worktime_3a.setObjectName("textEdit_worktime_3a")
            self.textEdit_worktime_3a.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_3a.setText(self.form_input_history["textEdit_worktime_3a"])

            self.textEdit_worktime_3b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_3b.setGeometry(QtCore.QRect(350, 510, 330, 40))
            self.textEdit_worktime_3b.setObjectName("textEdit_worktime_3b")
            self.textEdit_worktime_3b.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_3b.setText(self.form_input_history["textEdit_worktime_3b"])          
            
            self.textEdit_worktime_4a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_4a.setGeometry(QtCore.QRect(20, 540, 330, 40))
            self.textEdit_worktime_4a.setObjectName("textEdit_worktime_4a")
            self.textEdit_worktime_4a.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_4a.setText(self.form_input_history["textEdit_worktime_4a"])

            self.textEdit_worktime_4b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_4b.setGeometry(QtCore.QRect(350, 540, 330, 40))
            self.textEdit_worktime_4b.setObjectName("textEdit_worktime_4b")
            self.textEdit_worktime_4b.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_4b.setText(self.form_input_history["textEdit_worktime_4b"])

            self.textEdit_worktime_5a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_5a.setGeometry(QtCore.QRect(20, 570, 330, 40))
            self.textEdit_worktime_5a.setObjectName("textEdit_worktime_5a")
            self.textEdit_worktime_5a.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_5a.setText(self.form_input_history["textEdit_worktime_5a"])

            self.textEdit_worktime_5b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_5b.setGeometry(QtCore.QRect(350, 570, 330, 40))
            self.textEdit_worktime_5b.setObjectName("textEdit_worktime_5b")
            self.textEdit_worktime_5b.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_5b.setText(self.form_input_history["textEdit_worktime_5b"])

            self.textEdit_worktime_6a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_6a.setGeometry(QtCore.QRect(20, 600, 330, 40))
            self.textEdit_worktime_6a.setObjectName("textEdit_worktime_6a")
            self.textEdit_worktime_6a.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_6a.setText(self.form_input_history["textEdit_worktime_6a"])

            self.textEdit_worktime_6b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_6b.setGeometry(QtCore.QRect(350, 600, 330, 40))
            self.textEdit_worktime_6b.setObjectName("textEdit_worktime_6b")
            self.textEdit_worktime_6b.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_6b.setText(self.form_input_history["textEdit_worktime_6b"])

            self.textEdit_worktime_7a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_7a.setGeometry(QtCore.QRect(20, 630, 330, 40))
            self.textEdit_worktime_7a.setObjectName("textEdit_worktime_7a")
            self.textEdit_worktime_7a.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_7a.setText(self.form_input_history["textEdit_worktime_7a"])

            self.textEdit_worktime_7b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_7b.setGeometry(QtCore.QRect(350, 630, 330, 40))
            self.textEdit_worktime_7b.setObjectName("textEdit_worktime_7b")
            self.textEdit_worktime_7b.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_worktime_7b.setText(self.form_input_history["textEdit_worktime_7b"])
            
        
        else:
            self.comboBox_3 = self.createComboBox(self.tab, ["47N", "47S", "48N", "48S", "49N", "49S", "50S", "51S", "52S"], 540, 185)
            self.comboBox_3a = self.createComboBox(self.tab, ["< 10cm", "< 15cm", "< 20cm", "< 25cm", "< 30cm", "< 35cm",  "< 40cm"], 540, 210)
            self.lineEdit_2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_2.setGeometry(QtCore.QRect(210, 10, 320, 20))
            self.lineEdit_2.setObjectName("lineEdit_2")

            self.lineEdit_3 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_3.setGeometry(QtCore.QRect(210, 35, 320, 20))
            self.lineEdit_3.setObjectName("lineEdit_3")

            self.lineEdit_4 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_4.setGeometry(QtCore.QRect(210, 60, 320, 20))
            self.lineEdit_4.setObjectName("lineEdit_4")

            self.lineEdit_4a = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_4a.setGeometry(QtCore.QRect(540, 160, 140, 20))
            self.lineEdit_4a.setObjectName("lineEdit_4a")

            self.lineEdit_10a = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10a.setGeometry(QtCore.QRect(850, 10, 140, 20))
            self.lineEdit_10a.setObjectName("lineEdit_10a")

            self.lineEdit_10b = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10b.setGeometry(QtCore.QRect(850, 35, 140, 20))
            self.lineEdit_10b.setObjectName("lineEdit_10b")
            self.lineEdit_10b.setText("3 m")  # Set default value to "3 m"

            self.lineEdit_10c = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10c.setGeometry(QtCore.QRect(850, 60, 140, 20))
            self.lineEdit_10c.setObjectName("lineEdit_10c")
            self.lineEdit_10c.setText("2 m/s")  # Set default value to "2 m/s"

            self.lineEdit_10d = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10d.setGeometry(QtCore.QRect(850, 85, 140, 20))
            self.lineEdit_10d.setObjectName("lineEdit_10d")
            
            self.lineEdit_team = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_team.setGeometry(QtCore.QRect(850, 130, 140, 20))
            self.lineEdit_team.setObjectName("lineEdit_team") 
            
            self.lineEdit_combo6 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_combo6.setGeometry(QtCore.QRect(540, 135, 140, 20))
            self.lineEdit_combo6.setObjectName("lineEdit_combo6")
            
            self.lineEdit_10f = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10f.setGeometry(QtCore.QRect(850, 195, 140, 20))
            self.lineEdit_10f.setObjectName("lineEdit_10f")

            self.comboBox_10g = self.createComboBox(self.tab, ["Insecticide", "Herbicide", "Fertilizer", "Fungicide", "Larvacide"], 850, 220)

                
            self.lineEdit_10h = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10h.setGeometry(QtCore.QRect(850, 275, 140, 20))
            self.lineEdit_10h.setObjectName("lineEdit_10h")

            self.lineEdit_10i = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10i.setGeometry(QtCore.QRect(850, 420, 140, 20))
            self.lineEdit_10i.setObjectName("lineEdit_10i")

            self.lineEdit_10j = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10j.setGeometry(QtCore.QRect(850, 445, 140, 20))
            self.lineEdit_10j.setObjectName("lineEdit_10j")

            self.lineEdit_10k = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10k.setGeometry(QtCore.QRect(1180, 10, 140, 20))
            self.lineEdit_10k.setObjectName("lineEdit_10k")

            self.lineEdit_10l = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10l.setGeometry(QtCore.QRect(1180, 35, 140, 20))
            self.lineEdit_10l.setObjectName("lineEdit_10l")

            self.lineEdit_10m = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10m.setGeometry(QtCore.QRect(1180, 60, 140, 20))
            self.lineEdit_10m.setObjectName("lineEdit_10m")

            self.lineEdit_10n = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10n.setGeometry(QtCore.QRect(1180, 85, 140, 20))
            self.lineEdit_10n.setObjectName("lineEdit_10n")

            self.lineEdit_10o = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10o.setGeometry(QtCore.QRect(1180, 110, 140, 20))
            self.lineEdit_10o.setObjectName("lineEdit_10o")

            self.lineEdit_10p = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10p.setGeometry(QtCore.QRect(1180, 135, 140, 20))
            self.lineEdit_10p.setObjectName("lineEdit_10p")

            self.lineEdit_10q = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10q.setGeometry(QtCore.QRect(1180, 170, 140, 20))
            self.lineEdit_10q.setObjectName("lineEdit_10q")
            self.lineEdit_10q.setText("Cone")  # Set default value to "Cone"

            self.lineEdit_10r = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10r.setGeometry(QtCore.QRect(1180, 195, 140, 20))
            self.lineEdit_10r.setObjectName("lineEdit_10r")

            self.lineEdit_10s = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10s.setGeometry(QtCore.QRect(1180, 220, 140, 20))
            self.lineEdit_10s.setObjectName("lineEdit_10s")

            # Create and set up QSpinBox
            self.spinBox_10t = QtWidgets.QSpinBox(self.tab)
            self.spinBox_10t.setGeometry(QtCore.QRect(1180, 245, 35, 20))
            self.spinBox_10t.setFont(QtGui.QFont("Arial", 10))

            # Create and set up QComboBox
            self.comboBox_10t = QtWidgets.QComboBox(self.tab)
            self.comboBox_10t.setGeometry(QtCore.QRect(1225, 245, 95, 20))
            self.comboBox_10t.setFont(QtGui.QFont("Arial", 10))
            self.comboBox_10t.addItems(["Rotation", "New", "Spot Reservice"])


            self.lineEdit_10u = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10u.setGeometry(QtCore.QRect(1180, 305, 50, 20))
            self.lineEdit_10u.setObjectName("lineEdit_10u")
            self.lineEdit_10u.textChanged.connect(self.on_text_changed5)
            
            self.lineEdit_10u1 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10u1.setGeometry(QtCore.QRect(1235, 305, 40, 20))
            self.lineEdit_10u1.setObjectName("lineEdit_10u1")
            
            self.lineEdit_10u2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10u2.setGeometry(QtCore.QRect(1280, 305, 40, 20))
            self.lineEdit_10u2.setObjectName("lineEdit_10u2")

            self.lineEdit_10v = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10v.setGeometry(QtCore.QRect(1180, 330, 50, 20))
            self.lineEdit_10v.setObjectName("lineEdit_10v")
            self.lineEdit_10v.setText("**")  
            self.lineEdit_10v.textChanged.connect(self.on_text_changed1)
            
            self.lineEdit_10v1 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10v1.setGeometry(QtCore.QRect(1235, 330, 40, 20))
            self.lineEdit_10v1.setObjectName("lineEdit_10v1")
            self.lineEdit_10v1.setText("**")  
            
            self.lineEdit_10v2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10v2.setGeometry(QtCore.QRect(1280, 330, 40, 20))
            self.lineEdit_10v2.setObjectName("lineEdit_10v2")
            self.lineEdit_10v2.setText("**")  

            self.lineEdit_10w = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10w.setGeometry(QtCore.QRect(1180, 355, 50, 20))
            self.lineEdit_10w.setObjectName("lineEdit_10w")
            self.lineEdit_10w.setText("**")
            self.lineEdit_10w.textChanged.connect(self.on_text_changed2)
            
            self.lineEdit_10w1 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10w1.setGeometry(QtCore.QRect(1235, 355, 40, 20))
            self.lineEdit_10w1.setObjectName("lineEdit_10w1")
            self.lineEdit_10w1.setText("**")
            
            self.lineEdit_10w2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10w2.setGeometry(QtCore.QRect(1280, 355, 40, 20))
            self.lineEdit_10w2.setObjectName("lineEdit_10w2")
            self.lineEdit_10w2.setText("**")
            
            self.lineEdit_10x = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10x.setGeometry(QtCore.QRect(1180, 380, 50, 20))
            self.lineEdit_10x.setObjectName("lineEdit_10x")
            self.lineEdit_10x.setText("**")
            self.lineEdit_10x.textChanged.connect(self.on_text_changed3)

            self.lineEdit_10x1 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10x1.setGeometry(QtCore.QRect(1235, 380, 40, 20))
            self.lineEdit_10x1.setObjectName("lineEdit_10x1")
            self.lineEdit_10x1.setText("**")
            
            self.lineEdit_10x2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10x2.setGeometry(QtCore.QRect(1280, 380, 40, 20))
            self.lineEdit_10x2.setObjectName("lineEdit_10x2")
            self.lineEdit_10x2.setText("**")


            self.lineEdit_10y = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10y.setGeometry(QtCore.QRect(1180, 405, 50, 20))
            self.lineEdit_10y.setObjectName("lineEdit_10y")
            self.lineEdit_10y.setText("**")
            self.lineEdit_10y.textChanged.connect(self.on_text_changed4)
            
            self.lineEdit_10y1 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10y1.setGeometry(QtCore.QRect(1235, 405, 40, 20))
            self.lineEdit_10y1.setObjectName("lineEdit_10y1")
            self.lineEdit_10y1.setText("**")
            
            self.lineEdit_10y2 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10y2.setGeometry(QtCore.QRect(1280, 405, 40, 20))
            self.lineEdit_10y2.setObjectName("lineEdit_10y2")
            self.lineEdit_10y2.setText("**")

            self.comboBox_10z1 = self.createComboBox(self.tab, ["10 deg. (Gentle Slope)",  "20 deg.(Contour Slope)"], 1180, 445)

            self.lineEdit_10z3 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10z3.setGeometry(QtCore.QRect(850, 490, 330, 20))
            self.lineEdit_10z3.setObjectName("lineEdit_10z3")

            self.lineEdit_10z4 = QtWidgets.QLineEdit(self.tab)
            self.lineEdit_10z4.setGeometry(QtCore.QRect(850, 515, 330, 40))
            self.lineEdit_10z4.setObjectName("lineEdit_10z4")
            self.lineEdit_10z4.setText("Spraying Operation")  # Set default value to "3m"

            self.textEdit_10z5 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_10z5.setGeometry(QtCore.QRect(850, 560, 330, 40))
            self.textEdit_10z5.setObjectName("textEdit_10z5")
            self.textEdit_10z5.setWordWrapMode(QtGui.QTextOption.WordWrap)
            self.textEdit_10z5.setText("To complement the spraying with the ‘palm by palm’ method on plot number ")

            self.textEdit_situation_1 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_situation_1.setGeometry(QtCore.QRect(20, 270, 330, 40))
            self.textEdit_situation_1.setObjectName("textEdit_situation_1")
            self.textEdit_situation_1.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_solution_1 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_solution_1.setGeometry(QtCore.QRect(350, 270, 330, 40))
            self.textEdit_solution_1.setObjectName("textEdit_solution_1")
            self.textEdit_solution_1.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_situation_2 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_situation_2.setGeometry(QtCore.QRect(20, 310, 330, 40))
            self.textEdit_situation_2.setObjectName("textEdit_situation_2")
            self.textEdit_situation_2.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_solution_2 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_solution_2.setGeometry(QtCore.QRect(350, 310, 330, 40))
            self.textEdit_solution_2.setObjectName("textEdit_solution_2")
            self.textEdit_solution_2.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_situation_3 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_situation_3.setGeometry(QtCore.QRect(20, 350, 330, 40))
            self.textEdit_situation_3.setObjectName("textEdit_situation_3")
            self.textEdit_situation_3.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_solution_3 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_solution_3.setGeometry(QtCore.QRect(350, 350, 330, 40))
            self.textEdit_solution_3.setObjectName("textEdit_solution_3")
            self.textEdit_solution_3.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_situation_4 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_situation_4.setGeometry(QtCore.QRect(20, 390, 330, 40))
            self.textEdit_situation_4.setObjectName("textEdit_situation_4")
            self.textEdit_situation_4.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_solution_4 = QtWidgets.QTextEdit(self.tab)
            self.textEdit_solution_4.setGeometry(QtCore.QRect(350, 390, 330, 40))
            self.textEdit_solution_4.setObjectName("textEdit_solution_4")
            self.textEdit_solution_4.setWordWrapMode(QtGui.QTextOption.WordWrap)


            self.textEdit_worktime_1a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_1a.setGeometry(QtCore.QRect(20, 450, 330, 40))
            self.textEdit_worktime_1a.setObjectName("textEdit_worktime_1a")
            self.textEdit_worktime_1a.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_1b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_1b.setGeometry(QtCore.QRect(350, 450, 330, 40))
            self.textEdit_worktime_1b.setObjectName("textEdit_worktime_1b")
            self.textEdit_worktime_1b.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_2a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_2a.setGeometry(QtCore.QRect(20, 480, 330, 40))
            self.textEdit_worktime_2a.setObjectName("textEdit_worktime_2a")
            self.textEdit_worktime_2a.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_2b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_2b.setGeometry(QtCore.QRect(350, 480, 330, 40))
            self.textEdit_worktime_2b.setObjectName("textEdit_worktime_2b")
            self.textEdit_worktime_2b.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_3a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_3a.setGeometry(QtCore.QRect(20, 510, 330, 40))
            self.textEdit_worktime_3a.setObjectName("textEdit_worktime_3a")
            self.textEdit_worktime_3a.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_3b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_3b.setGeometry(QtCore.QRect(350, 510, 330, 40))
            self.textEdit_worktime_3b.setObjectName("textEdit_worktime_3b")
            self.textEdit_worktime_3b.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_4a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_4a.setGeometry(QtCore.QRect(20, 540, 330, 40))
            self.textEdit_worktime_4a.setObjectName("textEdit_worktime_4a")
            self.textEdit_worktime_4a.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_4b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_4b.setGeometry(QtCore.QRect(350, 540, 330, 40))
            self.textEdit_worktime_4b.setObjectName("textEdit_worktime_4b")
            self.textEdit_worktime_4b.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_5a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_5a.setGeometry(QtCore.QRect(20, 570, 330, 40))
            self.textEdit_worktime_5a.setObjectName("textEdit_worktime_5a")
            self.textEdit_worktime_5a.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_5b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_5b.setGeometry(QtCore.QRect(350, 570, 330, 40))
            self.textEdit_worktime_5b.setObjectName("textEdit_worktime_5b")
            self.textEdit_worktime_5b.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_6a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_6a.setGeometry(QtCore.QRect(20, 600, 330, 40))
            self.textEdit_worktime_6a.setObjectName("textEdit_worktime_6a")
            self.textEdit_worktime_6a.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_6b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_6b.setGeometry(QtCore.QRect(350, 600, 330, 40))
            self.textEdit_worktime_6b.setObjectName("textEdit_worktime_6b")
            self.textEdit_worktime_6b.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_7a = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_7a.setGeometry(QtCore.QRect(20, 630, 330, 40))
            self.textEdit_worktime_7a.setObjectName("textEdit_worktime_7a")
            self.textEdit_worktime_7a.setWordWrapMode(QtGui.QTextOption.WordWrap)

            self.textEdit_worktime_7b = QtWidgets.QTextEdit(self.tab)
            self.textEdit_worktime_7b.setGeometry(QtCore.QRect(350, 630, 330, 40))
            self.textEdit_worktime_7b.setObjectName("textEdit_worktime_7b")
            self.textEdit_worktime_7b.setWordWrapMode(QtGui.QTextOption.WordWrap)

        self.updateComboBox2()

    def updateComboBox2(self):
        selection = self.comboBox_2.currentText()
        if selection == "SMART Tbk.":
            items = ["Others","Pelakar","Tanah Laut", "Sungai Kupang", "Batu Ampar", "Kajui", "Jalemo", "Manuhing", "Naga Sakti", "Naga Mas", "Kijang", "Rama-Rama"]
        elif selection == "SD GUTHRIE":
            items = ["Others", "Chaah", "North Labis", "Sungai Simpang Kiri", "Yong Peng", "Kempas Klembang"]
        else:
            items = ["Others"]
        # Clear existing items and add new items
        self.comboBox_1.clear()
        self.comboBox_1.addItems(items)

    def on_text_changed1(self, text):
        self.lineEdit_10v.setText(text.upper())
    def on_text_changed2(self, text):
        self.lineEdit_10w.setText(text.upper())
    def on_text_changed3(self, text):
        self.lineEdit_10x.setText(text.upper())
    def on_text_changed4(self, text):
        self.lineEdit_10y.setText(text.upper())     
    def on_text_changed5(self, text):
        self.lineEdit_10u.setText(text.upper())   

    def setupMergeReportTab(self):
        self.MergePDF = QtWidgets.QWidget()
        self.MergePDF.setObjectName("MergePDF")

        labels = [
            ("label_11", "Location Folder CSV", 20, 20),
            ("label_12", "Location Folder Output", 20, 60),
            ("label_13", "Choose Company", 20, 100),
            ("label_14", "Choose Estate", 20, 140),
            ("label_15", "Name of Block", 20, 180),
            ("label_16", "Rotation Field", 20, 225),
            # ("label_17", "Date Start", 20, 265),
            # ("label_18", "Date End/Finish", 20, 307),
            ("label_21", "Area of Block (ha)", 20, 265),
            ("label_19", "*Please fill be carefull*", 20, 600),
            ("label_20", "Loading...", 1000, 590)
        ]

        for name, text, x, y in labels:
            label = QtWidgets.QLabel(self.MergePDF)
            label.setGeometry(QtCore.QRect(x, y, 180, 20))
            font = QtGui.QFont()
            font.setPointSize(10)
            label.setFont(font)
            label.setObjectName(name)
            label.setText(text)

        self.setupMergeReportTabInputs()
        self.tabWidget.addTab(self.MergePDF, "")

    def setupMergeReportTabInputs(self):
        self.pushButton_5 = QtWidgets.QPushButton(self.MergePDF)
        self.pushButton_5.setGeometry(QtCore.QRect(600, 20, 140, 25))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.pushButton_5.setFont(font)
        self.pushButton_5.setObjectName("pushButton_5")
        self.pushButton_5.setText("Folder .CSV")
        # self.pushButton_5.clicked.connect(lambda: self.selectFolder("Folder .PDF .CSV", self.lineEdit_5))  #Orinya
        self.pushButton_5.clicked.connect(lambda: self.pushButton_handler5("Folder .PDF .CSV"))

        self.pushButton_6 = QtWidgets.QPushButton(self.MergePDF)
        self.pushButton_6.setGeometry(QtCore.QRect(600, 60, 140, 25))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.pushButton_6.setFont(font)
        self.pushButton_6.setObjectName("pushButton_6")
        self.pushButton_6.setText("Folder Output")
        # self.pushButton_6.clicked.connect(lambda: self.selectFolder("Folder Output", self.lineEdit_6))   #Orinya
        self.pushButton_6.clicked.connect(lambda: self.pushButton_handler6("Folder Output"))

        self.pushButton_7 = QtWidgets.QPushButton(self.MergePDF)
        self.pushButton_7.setGeometry(QtCore.QRect(600, 420, 80, 25))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.pushButton_7.setFont(font)
        self.pushButton_7.setObjectName("pushButton_7")
        self.pushButton_7.setText("Process")
        self.pushButton_7.clicked.connect(lambda: self.pushButton_handler9("Process"))

        self.lineEdit_5 = QtWidgets.QLineEdit(self.MergePDF)
        self.lineEdit_5.setGeometry(QtCore.QRect(260, 20, 320, 25))
        self.lineEdit_5.setObjectName("lineEdit_5")

        self.lineEdit_6 = QtWidgets.QLineEdit(self.MergePDF)
        self.lineEdit_6.setGeometry(QtCore.QRect(260, 60, 320, 25))
        self.lineEdit_6.setObjectName("lineEdit_6")
        
        self.lineEdit_20 = QtWidgets.QLineEdit(self.MergePDF)
        self.lineEdit_20.setGeometry(QtCore.QRect(600, 180, 200, 25))
        self.lineEdit_20.setObjectName("lineEdit_20")

        self.comboBox_4a = self.createComboBox(self.MergePDF, ["SD GUTHRIE", "SMART Tbk","OTHERS"], 600, 100)
        self.comboBox_4 = self.createComboBox(self.MergePDF, ["CHAAH", "NORTH LABIS", "SUNGAI SIMPANG KIRI", "YONG PENG", "KEMPAS KLEMBANG", "KAJUI", "JALEMO", "MANUHING", "NAGA SAKTI", "NAGA MAS", "KIJANG", "RAMA-RAMA", "Pelakar","Tanah Laut", "Sungai Kupang", "Batu Ampar", "OTHERS"], 600, 140)

        self.spinBox_1 = QtWidgets.QSpinBox(self.MergePDF)
        self.spinBox_1.setGeometry(QtCore.QRect(600, 220, 40, 25))
        self.spinBox_1.setFont(QtGui.QFont("Arial", 10))
        self.spinBox_1.setObjectName("spinBox_1")
        
        self.lineEdit_21 = QtWidgets.QLineEdit(self.MergePDF)
        self.lineEdit_21.setGeometry(QtCore.QRect(600, 260, 200, 25))
        self.lineEdit_21.setObjectName("lineEdit_21")

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("Accuration Report_3.5", "Accuration Report_3.5"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("Accuration Report_3.5", "Daily Report"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.MergePDF), _translate("Accuration Report_3.5", "Merge Daily Report PDF"))
        
        self.comboBox_3.setItemText(0, _translate("Dialog", "47N"))
        self.comboBox_3.setItemText(1, _translate("Dialog", "47S"))
        self.comboBox_3.setItemText(2, _translate("Dialog", "48N"))
        self.comboBox_3.setItemText(3, _translate("Dialog", "48S"))
        self.comboBox_3.setItemText(4, _translate("Dialog", "49N"))
        self.comboBox_3.setItemText(5, _translate("Dialog", "49S"))
        self.comboBox_3.setItemText(6, _translate("Dialog", "50S"))
        self.comboBox_3.setItemText(7, _translate("Dialog", "51S"))
        self.comboBox_3.setItemText(8, _translate("Dialog", "52S"))
        
        self.comboBox_3a.setItemText(0, _translate("Dialog", "< 10cm"))
        self.comboBox_3a.setItemText(1, _translate("Dialog", "< 15cm"))
        self.comboBox_3a.setItemText(2, _translate("Dialog", "< 20cm"))
        self.comboBox_3a.setItemText(3, _translate("Dialog", "< 25cm"))
        self.comboBox_3a.setItemText(4, _translate("Dialog", "< 30cm"))
        self.comboBox_3a.setItemText(5, _translate("Dialog", "< 35cm"))
        self.comboBox_3a.setItemText(6, _translate("Dialog", "< 40cm"))
        

    def createComboBox(self, parent, items, x, y):
        comboBox = QtWidgets.QComboBox(parent)
        comboBox.setGeometry(QtCore.QRect(x, y, 140, 20))
        comboBox.setObjectName(f"comboBox_{x}_{y}")
        comboBox.addItems(items)
        return comboBox

    def createDateEdit(self, parent, x, y):
        dateEdit = QtWidgets.QDateEdit(parent)
        dateEdit.setGeometry(QtCore.QRect(x, y, 140, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        dateEdit.setFont(font)
        dateEdit.setDateTime(QtCore.QDateTime(QtCore.QDate(2024, 12, 1), QtCore.QTime(0, 0, 0)))
        dateEdit.setDate(QtCore.QDate(2024, 12, 1))
        dateEdit.setObjectName(f"DATE_{x}_{y}")
        return dateEdit

    def pushButton_handler2(self, button_name):
        self.data_log = QFileDialog.getExistingDirectory()
        self.lineEdit_2.setText(f"{button_name} - Directory Selected: {self.data_log}")

    def pushButton_handler3(self, button_name):
        self.data_waypoint = QFileDialog.getExistingDirectory()
        self.lineEdit_3.setText(f"{button_name} - Directory Selected: {self.data_waypoint}")

    def pushButton_handler4(self, button_name):
        self.data_output_report = QFileDialog.getExistingDirectory()
        self.lineEdit_4.setText(f"{button_name} - Directory Selected: {self.data_output_report}")

    def pushButton_handler5(self, button_name):
        self.data_pdf = QFileDialog.getExistingDirectory()
        self.lineEdit_5.setText(f"{button_name} - Directory Selected: {self.data_pdf}")

    def pushButton_handler6(self, button_name):
        self.data_output_pdfmerge = QFileDialog.getExistingDirectory()
        self.lineEdit_6.setText(f"{button_name} - Directory Selected: {self.data_output_pdfmerge}")


    def pushButton_handler1(self, action):
        if hasattr(self, 'data_log') and hasattr(self, 'data_waypoint') and hasattr(self, 'data_output_report'):
            if self.lineEdit_10u1.text().strip() and self.lineEdit_10u2.text().strip():
                
                self.process_data()
                QMessageBox.information(None, "PROCESS", "PROCESS DONE")
            else:
                QMessageBox.warning(None, "ALERT!!!", "Please fill in both SPH and Dose fields!")

        else :
            QMessageBox.warning(None, "ALERT!!!", "PLEASE FILL ALL THE PARAMETERS AND UTM ZONE!!! ")

   

    def pushButton_handler9(self, action):
        if hasattr(self, 'data_pdf') and hasattr(self, 'data_output_pdfmerge'):
            self.merge_data()
            QMessageBox.information(None, "PROCESS", "PROCESS DONE")
        else :
            QMessageBox.warning(None, "ALERT!!!", "PLEASE FILL ALL THE PARAMETERS!!! ")


    def process_data(self):
        try:
            
            self.progressBar.setValue(0)
            
            self.form_input_history["comboBox_3"] = self.comboBox_3.currentText()
            self.form_input_history["comboBox_3a"] = self.comboBox_3a.currentText()
            self.form_input_history["lineEdit_2"] = self.lineEdit_2.text()
            self.form_input_history["lineEdit_3"] = self.lineEdit_3.text()
            self.form_input_history["lineEdit_combo6"] = self.lineEdit_combo6.text()
            self.form_input_history["lineEdit_4"] = self.lineEdit_4.text()
            self.form_input_history["lineEdit_4a"] = self.lineEdit_4a.text()
            self.form_input_history["lineEdit_5"] = self.lineEdit_5.text()
            # self.form_input_history["lineEdit_6"] = lineEdit_6
            self.form_input_history["lineEdit_20"] = self.lineEdit_20.text()
            self.form_input_history["comboBox_4a"] = self.comboBox_4a.currentText() 
            self.form_input_history["comboBox_4"] = self.comboBox_4.currentText()
            self.form_input_history["lineEdit_21"] = self.lineEdit_21.text()
            self.form_input_history["lineEdit_10a"] = self.lineEdit_10a.text()
            self.form_input_history["lineEdit_10b"] = self.lineEdit_10b.text()
            self.form_input_history["lineEdit_10c"] = self.lineEdit_10c.text()
            self.form_input_history["lineEdit_10d"] = self.lineEdit_10d.text()
            self.form_input_history["lineEdit_team"] =self.lineEdit_team.text()
            self.form_input_history["lineEdit_10f"] = self.lineEdit_10f.text()
            self.form_input_history["comboBox_10g"] = self.comboBox_10g.currentText()
            self.form_input_history["lineEdit_10h"] = self.lineEdit_10h.text()
            self.form_input_history["lineEdit_10i"] = self.lineEdit_10i.text()
            self.form_input_history["lineEdit_10j"] = self.lineEdit_10j.text()
            self.form_input_history["lineEdit_10k"] = self.lineEdit_10k.text()
            self.form_input_history["lineEdit_10l"] = self.lineEdit_10l.text()
            self.form_input_history["lineEdit_10m"] = self.lineEdit_10m.text()
            self.form_input_history["lineEdit_10n"] = self.lineEdit_10n.text()
            self.form_input_history["lineEdit_10o"] = self.lineEdit_10o.text()
            self.form_input_history["lineEdit_10p"] = self.lineEdit_10p.text()
            self.form_input_history["lineEdit_10q"] = self.lineEdit_10q.text()
            self.form_input_history["lineEdit_10r"] = self.lineEdit_10r.text()
            self.form_input_history["lineEdit_10s"] = self.lineEdit_10s.text()
            self.form_input_history["comboBox_10t"] = self.comboBox_10t.currentText()
            self.form_input_history["comboBox_10z1"] = self.comboBox_10z1.currentText()
            self.form_input_history["lineEdit_10z3"] = self.lineEdit_10z3.text()
            self.form_input_history["lineEdit_10z4"] = self.lineEdit_10z4.text()
            
            
            self.form_input_history["textEdit_situation_1"] = self.textEdit_situation_1.toPlainText()
            self.form_input_history["textEdit_situation_2"] = self.textEdit_situation_2.toPlainText()
            self.form_input_history["textEdit_situation_3"] = self.textEdit_situation_3.toPlainText()
            self.form_input_history["textEdit_situation_4"] = self.textEdit_situation_4.toPlainText()
            self.form_input_history["textEdit_solution_1"] = self.textEdit_solution_1.toPlainText()
            self.form_input_history["textEdit_solution_2"] = self.textEdit_solution_2.toPlainText()
            self.form_input_history["textEdit_solution_3"] = self.textEdit_solution_3.toPlainText()
            self.form_input_history["textEdit_solution_4"] = self.textEdit_solution_4.toPlainText()
            
            
            self.form_input_history["textEdit_worktime_1a"] = self.textEdit_worktime_1a.toPlainText()
            self.form_input_history["textEdit_worktime_1b"] = self.textEdit_worktime_1b.toPlainText()
            self.form_input_history["textEdit_worktime_2a"] = self.textEdit_worktime_2a.toPlainText()
            self.form_input_history["textEdit_worktime_2b"] = self.textEdit_worktime_2b.toPlainText()
            self.form_input_history["textEdit_worktime_3a"] = self.textEdit_worktime_3a.toPlainText()
            self.form_input_history["textEdit_worktime_3b"] = self.textEdit_worktime_3b.toPlainText()
            self.form_input_history["textEdit_worktime_4a"] = self.textEdit_worktime_4a.toPlainText()
            self.form_input_history["textEdit_worktime_4b"] = self.textEdit_worktime_4b.toPlainText()
            self.form_input_history["textEdit_worktime_5a"] = self.textEdit_worktime_5a.toPlainText()
            self.form_input_history["textEdit_worktime_5b"] = self.textEdit_worktime_5b.toPlainText()
            self.form_input_history["textEdit_worktime_6a"] = self.textEdit_worktime_6a.toPlainText()
            self.form_input_history["textEdit_worktime_6b"] = self.textEdit_worktime_6b.toPlainText()
            self.form_input_history["textEdit_worktime_7a"] = self.textEdit_worktime_7a.toPlainText()
            self.form_input_history["textEdit_worktime_7b"] = self.textEdit_worktime_7b.toPlainText()

            
            
            self.save_form_input_history()

            print(f"Lokasi folder log adalah {self.data_log}")
            print(f"Lokasi folder waypoint adalah {self.data_waypoint}")
            print(f"Lokasi folder output adalah {self.data_output_report}")
            print(f"Lokasi folder output adalah {self.data_output_report}")
            print(f"Lokasi Estate adalah {self.comboBox_1.currentText()}")
            print(f"Lokasi Petak adalah {self.lineEdit_10u.text()}")
            print(f"Drone ID adalah {self.lineEdit_4a.text()}")
            print(f"Lokasi UTM Adalah {self.comboBox_3.currentText()}")
            # print(f"Tanggal {self.DATE_1.date().toString('dd-MMMM-yyyy')}")
            print(f"TInggak Akurasi yang di pilih adalah {self.comboBox_3a.currentText()} cm")
            
            # e = self.DATE_1.date().toString('dd-MMMM-yyyy')
            list_directory = ["junk" ,f"result {self.lineEdit_4a.text()}_{self.comboBox_1.currentText()}_{self.comboBox_2.currentText()}"]
            junk = os.path.join(self.data_output_report,list_directory[0])
            result = os.path.join(self.data_output_report,list_directory[1])
            
            os.makedirs(junk, exist_ok=True)
            os.makedirs(result, exist_ok=True)
            
            pie = os.path.join(junk, "pie")
            plot = os.path.join(junk, "plot")
            temporary = os.path.join(junk, "temporary")
            wp = os.path.join(junk, "wp")
            log = os.path.join(junk, "log")
            utm_flight = os.path.join(junk, "utm_flight")
            path_wp = os.path.join(junk, "path_wp")
            
            os.makedirs(pie, exist_ok=True)
            os.makedirs(plot, exist_ok=True)
            os.makedirs(temporary, exist_ok=True)
            os.makedirs(wp, exist_ok=True)
            os.makedirs(utm_flight, exist_ok=True)
            os.makedirs(log, exist_ok=True)
            os.makedirs(path_wp, exist_ok=True )
            

            zone_mapping = {
                "47N": 'EPSG:32647',
                "47S": 'EPSG:32747',
                "48N": 'EPSG:32648',
                "48S": 'EPSG:32748',
                "49N": 'EPSG:32649',
                "49S": 'EPSG:32749',
                "50S": 'EPSG:32750',
                "51S": 'EPSG:32751',
                "52S": 'EPSG:32752'
            }
            
            accuracy = {
                "< 10cm" : 0.1,
                "< 15cm" : 0.15,
                "< 20cm" : 0.2,
                "< 25cm" : 0.25,
                "< 30cm" : 0.3,
                "< 35cm" : 0.35,
                "< 40cm" : 0.4
            }
            
            def convert_value(value):
                return int(value) if value != "**" else value
            
            blok1 = self.lineEdit_10u.text()
            sph1 = convert_value(self.lineEdit_10u1.text())
            dosis1 = convert_value(self.lineEdit_10u2.text())
            
            blok2 = self.lineEdit_10v.text()
            sph2 = convert_value(self.lineEdit_10v1.text())
            dosis2 = convert_value(self.lineEdit_10v2.text())
            
            blok3 = self.lineEdit_10w.text()
            sph3 = convert_value(self.lineEdit_10w1.text())
            dosis3 = convert_value(self.lineEdit_10w2.text())
            
            blok4 = self.lineEdit_10x.text()
            sph4 = convert_value(self.lineEdit_10x1.text())
            dosis4 = convert_value(self.lineEdit_10x2.text())
            
            blok5 = self.lineEdit_10y.text()
            sph5 = convert_value(self.lineEdit_10y1.text())
            dosis5 = convert_value(self.lineEdit_10y2.text())
            
            
            block_all =[]
            
            blok_semua = [blok1,blok2,blok3,blok4,blok5]
            for block_ in blok_semua:
                if block_ != "**":
                    block_all.append(block_)
                    
            block_all = ', '.join(block_all)
            print(block_all)
            
            
            
            class PDF(FPDF):
                def add_dynamic_header(sel, company, estate, petak, date, drone_id, accu ):
                    sel.header()
                    sel.set_font('Arial', 'B', 14)
                    sel.cell(0, 5, 'ACCURACY STATUS REPORT', 0, 1, 'C')
                    
                    sel.set_font('Arial', '', 12)
                    sel.cell(0, 5, company, 0, 1, 'C')

                    sel.set_font('Arial', '', 12)
                    sel.cell(0, 5, estate, 0, 1, 'C')

                    sel.set_font('Arial', '', 12)
                    sel.cell(0, 5, petak, 0, 1, 'C')

                    sel.set_font('Arial', '', 12)
                    sel.cell(0, 5, date, 0, 1, 'C')

                    sel.set_font('Arial', '', 12)
                    sel.cell(0, 5, drone_id , 0, 1, 'C')
                    
                    sel.set_font('Arial', '', 12)
                    sel.cell(0, 5,  f"Accuration {accu}" , 0, 1, 'C')
                
                def summary(sel,company1, estate1, drone_id1, date1, accu):
                    sel.header()
                    sel.set_font('Arial', 'B', 14)
                    sel.cell(0, 5, 'SUMMARY ACCURACY STATUS REPORT', 0, 1, 'C')
                    
                    sel.set_font('Arial', '', 12)
                    sel.cell(0, 5, company1, 0, 1, 'C')

                    sel.set_font('Arial', '', 12)
                    sel.cell(0, 5, estate1, 0, 1, 'C')

                    sel.set_font('Arial', '', 12)
                    sel.cell(0, 5, drone_id1, 0, 1, 'C')

                    sel.set_font('Arial', '', 12)
                    sel.cell(0, 5, date1, 0, 1, 'C')
                    
                    sel.set_font('Arial', '', 12)
                    sel.cell(0, 5, f"Accuration {accu}" , 0, 1, 'C')
                    
                def footer(self):
                    self.set_y(-15)
                    self.set_font('Arial', 'I', 8)
                    self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

                def add_image(self, image_path, x, y, w, h):
                    self.image(image_path, x, y, w, h)
                    
                def add_table(sel, dataframe,x ,y, table_title):
                    sel.set_xy(x-20, y)
                    sel.set_font('Arial', 'B', 12)
                    sel.cell(0, 10, table_title, 0, 1, 'C')
                    sel.set_xy(x, y + 10)
                    sel.set_font('Arial', 'B', 12)
                    sel.cell(60, 8, 'Status', 1, 0, 'C')
                    sel.cell(40, 8, 'Number of Trees', 1, 0, 'C')
                    sel.cell(40, 8, 'Percentage', 1, 0, 'C')
                    sel.ln()
                    sel.set_font('Arial', '', 10)
                    for i in range(len(dataframe)):
                        sel.set_x(x)
                        if i == len(dataframe) - 1:  # Check if this is the last row (Total row)
                            sel.set_font('Arial', 'B', 10)  # Set font to bold for the Total row
                        else:
                            sel.set_font('Arial', '', 10)  # Set font to regular for other rosheet_df
                        sel.set_fill_color(200, 220, 255)
                        fill = True
                        sel.cell(60, 6, str(dataframe.iloc[i, 0]), 1, 0, 'C')
                        sel.cell(40, 6, str(dataframe.iloc[i, 1]), 1, 0, 'C')
                        sel.set_fill_color(200, 220, 255)
                        sel.cell(40, 6, f"{dataframe.iloc[i, 2]}", 1, 0,'C',fill)
                        sel.ln()

            pdf = PDF()
            pdf = PDF(format='A4')
            pdf.set_auto_page_break(auto=True, margin=15)
            
            wp_files = [file for file in os.listdir(self.data_waypoint) if file.endswith(".waypoints") or file.endswith(".plan") ]
            if not wp_files:
                QMessageBox.warning(None, "ALERT!!!", "SELECT A FOLDER CONTAINS .WAYPOINTS FILE ")
                return
                    
            log_files = [file for file in os.listdir(self.data_log) if file.endswith(".bin") or file.endswith(".BIN")]
            if not log_files:
                QMessageBox.warning(None, "ALERT!!!", "SELECT A FOLDER CONTAINS .bin or .BIN FILE ")
                return  
            
            for file in wp_files:
                if file.find(".waypoints") != -1:
                    base = os.path.splitext(file)[0]
                    with open(os.path.join(self.data_waypoint,file), 'r') as file:
                        filtered_data = []
                        lines = file.readlines()
                        try:
                            for line in lines[1:]:
                                if line.strip():  # Mengabaikan baris kosong
                                    parts = line.strip().split('\t')
                                    # Memastikan ada cukup kolom sebelum mengambil nilai ke-8 dan ke-9
                                    data_sebelum_lat_long = parts[7]  # Data sebelum latitude dan longitude
                                    latitude = parts[8]
                                    longitude = parts[9]
                                    if float(data_sebelum_lat_long) > 0:  # Misalnya kondisi data sebelum lat/long harus lebih besar dari 0
                                        filtered_data.append({
                                            'x': float(longitude),
                                            'y': float(latitude),
                                            'x1' : float(longitude),
                                            'y1' : float(latitude),
                                            'misi' : str(base)
                                        })
                                    df = pd.DataFrame(filtered_data, columns=['x', 'y', 'x1', 'y1', 'misi'])
                        except:
                            pass
                        
                        geometry = [Point(xy) for xy in zip(df['x'], df['y'])]
                        crs_wgs84 = 'EPSG:4326'
                        crs_utm48n = zone_mapping[self.comboBox_3.currentText()]
                        
                        transformer = pyproj.Transformer.from_crs(crs_wgs84, crs_utm48n, always_xy=True)
                        transformed_geometry = [Point(transformer.transform(point.x, point.y)) for point in geometry]
                        point_wp_proj = gpd.GeoDataFrame(df, geometry=transformed_geometry, crs=crs_utm48n)

                        point_wp_proj['x'] = point_wp_proj['geometry'].x
                        point_wp_proj['y'] = point_wp_proj['geometry'].y
                        # point_wp_proj['tgl'] = tanggal
                        data_pohon = point_wp_proj
                        data_pohon.to_csv(os.path.join(wp, base + ".csv"))
                
                elif file.find(".plan") != -1:
                    base = os.path.splitext(file)[0]
                    with open(os.path.join(self.data_waypoint,file), 'r') as file:
                        data = json.load(file)
                        coordinates = []
                        try:
                            items_184 = [item for item in data.get("mission", {}).get("items", []) if item.get("command") == 184]
                            for item in items_184:
                                lat = item.get("params")[4]
                                lon = item.get("params")[5]
                                if lat is not None and lon is not None:
                                    coordinates.append({
                                                'x': float(lon),
                                                'y': float(lat),
                                                'x1' : float(lon),
                                                'y1' : float(lat),
                                                'misi' : str(base)
                                            })
                                df = pd.DataFrame(coordinates, columns=['x', 'y', 'x1', 'y1', 'misi'])
                        except:
                            pass
                        
                            
                        geometry = [Point(xy) for xy in zip(df['x'], df['y'])]
                        crs_wgs84 = 'EPSG:4326'
                        crs_utm48n = zone_mapping[self.comboBox_3.currentText()]
                        
                        transformer = pyproj.Transformer.from_crs(crs_wgs84, crs_utm48n, always_xy=True)
                        transformed_geometry = [Point(transformer.transform(point.x, point.y)) for point in geometry]
                        point_wp_proj = gpd.GeoDataFrame(df, geometry=transformed_geometry, crs=crs_utm48n)

                        point_wp_proj['x'] = point_wp_proj['geometry'].x
                        point_wp_proj['y'] = point_wp_proj['geometry'].y
                        # point_wp_proj['tgl'] = tanggal
                        data_pohon = point_wp_proj
                        data_pohon.to_csv(os.path.join(wp, base + ".csv"))

            
            self.progressBar.setValue(10)
            
            gps_epoch = datetime(1980, 1, 6)
                
            all_data2 = pd.DataFrame()

            for file in log_files:
                base_name = os.path.splitext(file)[0]  # Ambil nama file tanpa ekstensi

                # Buka file .bin
                log_file_path = os.path.join(self.data_log, file)
                log_file = mavutil.mavlink_connection(log_file_path, robust_parsing=True)

                gps_all = []

                # Loop untuk membaca pesan dari file
                while True:
                    msg = log_file.recv_match(blocking=True)
                    if msg is None:
                        break

                    msg_dict = msg.to_dict()

                    # Ambil hanya pesan tipe 'GPS' dengan 'U' bernilai 1
                    if msg.get_type() == 'GPS' and msg_dict.get('U') == 1:
                        gps_all.append(msg_dict)

                # Siapkan data untuk disimpan ke CSV
                longitude = []
                latitude = []
                week = []
                waktu = []

                for row in gps_all:
                    longitude.append(row['Lng'])
                    latitude.append(row['Lat'])
                    week.append(row['GWk'])
                    waktu.append(row['GMS'])

                # Buat DataFrame
                data_dict = {"x": longitude, "y": latitude, "week": week, "waktu": waktu}
                df = pd.DataFrame(data_dict)
                df.to_csv(os.path.join(log, base_name + ".csv"))

                # Fungsi konversi GPS ke UTC
                def convert_gps_to_utc(row):
                    # Setiap minggu memiliki 7 hari
                    days_since_epoch = row['week'] * 7
                    # Hitung tanggal dengan menambahkan hari
                    target_date = gps_epoch + timedelta(days=days_since_epoch)
                    # Konversi milidetik ke detik
                    seconds = row['waktu'] / 1000.0
                    # Tambahkan detik ke tanggal target
                    target_datetime = target_date + timedelta(seconds=seconds)
                    # Konversi ke UTC+7
                    utc_plus_7_time = target_datetime + timedelta(hours=7)                
                    return utc_plus_7_time

                # Konversi waktu GPS ke UTC
                df['waktu'] = df.apply(convert_gps_to_utc, axis=1)

                # Hitung durasi, waktu mulai, dan tanggal
                start = df['waktu'].iloc[0]
                end = df['waktu'].max()
                duration = end - start
                date1 = start.date()

                start_time = start.strftime("%H:%M:%S")
                duration_datetime = datetime(1, 1, 1) + duration
                formatted_duration = duration_datetime.strftime("%H:%M:%S")

                start1 = [start_time]
                selisih = [formatted_duration]
                tgl = [date1]

                temp_df = pd.DataFrame({"start": start1, "selisih": selisih, 'tanggal': tgl})

                # Gabungkan DataFrame temp_df dengan all_data2
                all_data2 = pd.concat([all_data2, temp_df], ignore_index=True)

            # Format tanggal
            all_data2['tanggal'] = pd.to_datetime(all_data2['tanggal']).dt.strftime('%d-%B-%Y')

            # Simpan all_data2 ke CSV
            all_data2.to_csv(os.path.join(junk, 'tabel1.csv'), index=True)

            self.progressBar.setValue(20)  
            tanggal1 = all_data2['tanggal'].mode()[0]

            
            csv_files = [file for file in os.listdir(log) if file.endswith('.csv')]
            dfs = []
            for csv_file in csv_files:
                df = pd.read_csv(os.path.join(log, csv_file))
                dfs.append(df)
            merged_df = pd.concat(dfs, ignore_index=True)
            merged_df.to_csv(os.path.join(log, 'merged_gps.csv'), index=False)
            # print("Data GPS dari semua file log telah digabungkan dalam satu file CSV.") 

            point_uav = pd.read_csv(os.path.join(log, 'merged_gps.csv'))
            point_uav = point_uav.drop(columns='Unnamed: 0')
            point_uav ['x2'] = point_uav['x'].copy()
            point_uav ['y2'] = point_uav['y'].copy()

            geometry = [Point(xy) for xy in zip(point_uav['x'], point_uav['y'])]
            crs_wgs84 = 'EPSG:4326'
            crs_utm48n = zone_mapping[self.comboBox_3.currentText()]

            transformer = pyproj.Transformer.from_crs(crs_wgs84, crs_utm48n, always_xy=True)
            transformed_geometry = [Point(transformer.transform(point.x, point.y)) for point in geometry]
            point_uav_proj = gpd.GeoDataFrame(point_uav, geometry=transformed_geometry, crs=crs_utm48n)
            point_uav_proj['no'] = range(1, len(point_uav_proj) + 1)

            point_uav_proj['x'] = point_uav_proj['geometry'].x
            point_uav_proj['y'] = point_uav_proj['geometry'].y
            data_uav = point_uav_proj
            data_uav.to_csv(os.path.join(log, 'merged_gps_all.csv'))
            
            self.progressBar.setValue(30)
            # print("oke3")
            all_data = pd.DataFrame(columns=['nama', 'total'])
            index1 = 1
            for file_name in sorted(os.listdir(wp)):
                if file_name.endswith('.csv'):
                    base1 = os.path.splitext(file_name)[0]
                    file_path = os.path.join(wp, file_name)
                    data_pohon_df = pd.read_csv(file_path)
                    
                    x = data_pohon_df['x']
                    y = data_pohon_df['y']
                    plt.plot(x, y, marker='o', markerfacecolor='0', color='r', zorder=1)
                    for i in range(len(x)):
                        plt.text(x[i], y[i], str(i+1), fontsize=8, ha='right')
                    plt.xlabel('Longitude')
                    plt.ylabel('Latitude')
                    plt.title(base1)
                    plt.grid(True)
                    path_ = os.path.join(path_wp, f'{index1}.png')
                    plt.savefig(path_, format='png', bbox_inches='tight', dpi=200)
                    plt.close()
                                    
                    data_pohon_df['geometry'] = data_pohon_df['geometry'].apply(loads)        
                    data_pohon = gpd.GeoDataFrame(data_pohon_df, geometry='geometry')

                    buffer_pohon = data_pohon.buffer(4)
                    hasil_pengukuran_terdekat = gpd.GeoDataFrame(columns=['x_penyemprotan', 'y_penyemprotan', 'jarak', 'Status', 'x1', 'y1', 'x2', 'y2', 'misi'])
                                
                    for idx, pohon in data_pohon.iterrows():
                        buffer_area = buffer_pohon.iloc[idx]
                        data_terbang_dalam_buffer = data_uav[data_uav.geometry.intersects(buffer_area)]

                        titik_terbang_terdekat = None
                        jarak_terpendek = float('inf')
                        for idx, terbang in data_terbang_dalam_buffer.iterrows():
                            jarak = pohon.geometry.distance(terbang.geometry)
                            if jarak < jarak_terpendek:
                                jarak_terpendek = jarak
                                titik_terbang_terdekat = terbang.geometry

                        selected_accuracy = accuracy[self.comboBox_3a.currentText()]
                        
                        if jarak_terpendek < selected_accuracy:
                            status_penyemprotan = 'Accurate'
                        elif jarak_terpendek <= 1:  # Antara akurasi dan 1 meter
                            status_penyemprotan = 'Not Accurate'
                        elif jarak_terpendek > 1 and jarak_terpendek <= 4:
                            status_penyemprotan = 'Not Sprayed'
                        else:  # Lebih dari 4 meter
                            status_penyemprotan = 'Not Sprayed'
                            
                        
                        if titik_terbang_terdekat is not None:
                            hasil_pengukuran_terdekat = hasil_pengukuran_terdekat._append({'x_penyemprotan': titik_terbang_terdekat.coords[0][0],
                                                                                            'y_penyemprotan': titik_terbang_terdekat.coords[0][1], 
                                                                                            'jarak': jarak_terpendek, 
                                                                                            'Status': status_penyemprotan, 
                                                                                            'x1' : pohon['x1'], 
                                                                                            'y1' : pohon['y1'], 
                                                                                            'x2': titik_terbang_terdekat.coords[0][0], 
                                                                                            'y2' : titik_terbang_terdekat.coords[0][1], 
                                                                                            'misi': pohon['misi']}, ignore_index=True)

                        else:
                            hasil_pengukuran_terdekat = hasil_pengukuran_terdekat._append({'x_pohon': pohon['x'], 
                                                                                            'y_pohon': pohon['y'], 
                                                                                            'Status': status_penyemprotan, 
                                                                                            'x1' : pohon['x1'], 
                                                                                            'y1' : pohon['y1'], 
                                                                                            'misi': pohon['misi']}, ignore_index=True)
                                                
                    hasil_pengukuran_terdekat.to_csv(os.path.join(temporary, base1 + ".csv"))
                    
                    # print(terbang)
                    
                    contoh_pie = hasil_pengukuran_terdekat[hasil_pengukuran_terdekat['Status'].isin(['Accurate', 'Not Accurate'])]
                    data_counts1 = contoh_pie['Status'].value_counts()
                    
                    data_counts = hasil_pengukuran_terdekat['Status'].value_counts()

                    status_color_map = {
                    'Accurate': 'limegreen',
                    'Not Accurate': 'dodgerblue',
                    # 'Missed': 'red',
                    'Not Sprayed' : 'brown'
                    }

                    # Pastikan urutan warna sesuai dengan urutan status di pie chart
                    colors = [status_color_map[status] for status in data_counts1.index]

                    # Visualisasi pie chart
                    explode = [0.02] * len(data_counts1)
                    plt.pie(data_counts1.values, labels=data_counts1.index, autopct='%0.01f%%', 
                        explode=explode, shadow={'ox': -0.01, 'edgecolor': 'none', 'shade': 0.9}, 
                        startangle=90, colors=colors)

                    plt.title('Summary Status', fontweight='bold')
                    pie_chart = os.path.join(pie, f'{index1}.png')
                    plt.savefig(pie_chart, format='png', bbox_inches='tight', dpi=200)
                    plt.close()

                    
                    x_coords = hasil_pengukuran_terdekat['x1'] 
                    y_coords= hasil_pengukuran_terdekat['y1']
                    status = hasil_pengukuran_terdekat['Status']
                    colors = {'Accurate': 'limegreen', 
                            #   'Missed': 'red', 
                                'Not Accurate': 'dodgerblue',
                                'Not Sprayed' : 'brown'}
                    
                    sns.scatterplot(x=x_coords, y=y_coords, hue=status, palette=colors, data=hasil_pengukuran_terdekat, s=100)
                    plt.xlabel('Longitude')
                    plt.ylabel('Latitude')
                    plt.title('Status of Tree')
                    plt.grid(True)
                    plt.gca().xaxis.get_major_formatter().set_useOffset(False)
                    # plot_ = os.path.join(plot + '3.png')
                    plot_ = os.path.join(plot, f'{index1}.png')
                    plt.savefig(plot_, format='png', bbox_inches='tight', dpi=200)
                    plt.close()
                    
                    
                                        
                    df1 = pd.DataFrame()
                    df1['Status'] = data_counts.index
                    df1['Number of Trees'] = data_counts.values

                    total_sprayed = df1[df1['Status'].isin(['Accurate', 'Not Accurate'])]['Number of Trees'].sum()
                    total_mission = total_sprayed + df1[df1['Status'] == 'Not Sprayed']['Number of Trees'].sum()

                    # Menghitung persentase untuk Accurate dan Not Accurate
                    df1.loc[df1['Status'] == 'Accurate', 'Percentage'] = ((df1.loc[df1['Status'] == 'Accurate', 'Number of Trees'] / total_sprayed * 100).round(2).astype(str) + "%")
                    df1.loc[df1['Status'] == 'Not Accurate', 'Percentage'] = ((df1.loc[df1['Status'] == 'Not Accurate', 'Number of Trees'] / total_sprayed * 100).round(2).astype(str) + "%")
                    df1.loc[df1['Status'] == 'Not Sprayed', 'Percentage'] = ""
                            
                    df1 = pd.concat([df1, pd.DataFrame([
                        ['Total Sprayed', total_sprayed, '100 %'],
                        ['Total Mission', total_mission, '']
                    ], columns=['Status', 'Number of Trees', 'Percentage'])], ignore_index=True)
                    
                    
                    status_order = ["Accurate", "Not Accurate", "Not Sprayed", "Total Sprayed", "Total Mission"]
                    df1['Status'] = pd.Categorical(df1['Status'], categories=status_order, ordered=True)
                    df1 = df1.sort_values('Status').reset_index(drop=True)
                    print(df1)



                    a = [total_sprayed]
                    temp_df = pd.DataFrame({"nama": base1, "total":a})
                    
                    # Tambahkan data sementara ke DataFrame utama
                    all_data = pd.concat([all_data, temp_df], ignore_index=True)
                    # print(all_data)
                    
                    pdf.add_page()
                    company  = f'{self.comboBox_2.currentText()}'
                    estate = f'{self.comboBox_1.currentText()}'
                    date = f'{tanggal1}'
                    drone_id = f'{self.lineEdit_4a.text()}'
                    petak = block_all
                    accu = f'{self.comboBox_3a.currentText()}'
                    pdf.add_dynamic_header(company, estate, petak, date, drone_id, accu)

                    # Menambahkan gambar ke PDF
                    pdf.add_image(path_, 10, 50, 88, 75)
                    pdf.add_image(plot_, 105, 50, 96, 65)
                    pdf.add_image(pie_chart, 75, 140, 70, 70)
                    pdf.add_table(df1, 35, 220, 'Spray Accuration Summary')

                    index1 += 1
            self.progressBar.setValue(60)
            all_data.to_csv(os.path.join(junk, 'tabel.csv'), index=True)
            
            df_merged = pd.concat([all_data, all_data2], axis=1)
            df_merged['tanggal'] = pd.to_datetime(df_merged['tanggal']).dt.strftime('%d-%B-%Y')
            df_merged.to_excel(os.path.join(junk, 'GABUNG.xlsx'), index=False)        
            # tanggal = df_merged['tanggal'].iloc[-1]
            tanggal= df_merged['tanggal'].mode()[0]

            print(tanggal)
            # tanggal = pd.to_datetime(tanggal).strftime('%d-%B-%Y')

            csv_final = [file for file in os.listdir(temporary) if file.endswith('.csv')]
            merge_final = []
            for csv_merge in csv_final:
                df = pd.read_csv(os.path.join(temporary, csv_merge))
                merge_final.append(df)
            merged_csv_final = pd.concat(merge_final, ignore_index=True)
            
            
            df_merged1 = pd.DataFrame({'tanggal': [df_merged['tanggal'].iloc[0]]})
            df1_repeated = pd.concat([df_merged1]*len(merged_csv_final), ignore_index=True)
            merged_csv_final = pd.concat([merged_csv_final, df1_repeated], axis=1)
            
            
            crs_wgs84 = 'EPSG:4326'
            crs_utm48n = zone_mapping[self.comboBox_3.currentText()]
            transformer = pyproj.Transformer.from_crs(crs_utm48n, crs_wgs84, always_xy=True)

            # Inisialisasi list untuk menyimpan geometry
            transformed_geometry = []

            # Iterasi per baris dengan pengecekan
            for index, row in merged_csv_final.iterrows():
                if pd.notna(row['x2']) and pd.notna(row['y2']):
                    try:
                        # Transformasi hanya untuk baris dengan data valid
                        x, y = transformer.transform(row['x2'], row['y2'])
                        transformed_geometry.append(Point(x, y))
                    except Exception as e:
                        print(f"Error pada baris {index}: {e}")
                        transformed_geometry.append(None)
                else:
                    # Jika data kosong, tambahkan None
                    transformed_geometry.append(None)

            # Tambahkan kolom geometry ke dataframe asli
            merged_csv_final['geometry'] = transformed_geometry

            # Buat GeoDataFrame
            point_uav_final = gpd.GeoDataFrame(merged_csv_final, geometry='geometry', crs=crs_wgs84)

            # Perbarui kolom x2 dan y2 hanya untuk baris valid
            point_uav_final['x2'] = point_uav_final['geometry'].apply(lambda geom: geom.x if geom else None)
            point_uav_final['y2'] = point_uav_final['geometry'].apply(lambda geom: geom.y if geom else None)

            # Final dataframe tetap dalam urutan aslinya
            merged_csv_final1 = point_uav_final

            
            merged_csv_final1.to_csv(os.path.join (result, f"Daily Accuration Report {self.lineEdit_4a.text()} {self.comboBox_1.currentText()} {block_all} {tanggal} .csv"), index=False)
            
            hti = Html2Image(custom_flags=['--headless', '--disable-gpu'])
            timeout = 5
            
            plot_final = os.path.join(plot, 'final.png')
            try:
                requests.head("http://www.google.com/", timeout=timeout)
                # Do something

                # Menghitung pusat dari semua titik (centroid) untuk peta
                center_lat = merged_csv_final['y1'].mean()
                center_lon = merged_csv_final['x1'].mean()

                # Membuat peta dengan pusat yang dihitung
                m = folium.Map(location=[center_lat, center_lon], zoom_start=17.2)

                tile_url = 'https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}'
                folium.TileLayer(
                    tiles=tile_url,
                    attr='Google',
                    name='Google Satellite',
                    overlay=False,
                    control=True
                ).add_to(m)

                # Warna yang digunakan berdasarkan status
                colors = {'Accurate': 'limegreen', 
                        #   'Missed': 'red', 
                            'Not Accurate': 'dodgerblue', 
                            'Not Sprayed' : 'brown'}

                # Menambahkan titik-titik dengan warna sesuai status
                for _, row in merged_csv_final.iterrows():
                    coord = (row['y1'], row['x1'])  # lat (y_penyemprotan), lon (x_penyemprotan)
                    color = colors.get(row['Status'], 'black')  # Gunakan warna sesuai status, atau default 'black'
                    
                    folium.CircleMarker(
                        location=coord,
                        radius=3,  # Ukuran marker
                        color=color,
                        fill=True,
                        fill_opacity=1
                    ).add_to(m)
                    
                # plot_final = os.path.join(plot, 'final.png')
                # plt.savefig(plot_final, format='png', bbox_inches='tight', dpi=200)
                
                # Menyimpan peta ke file HTML
                m.save(os.path.join(plot, 'final.html'))
                # print('The internet connection is active')

                # Menggunakan Selenium untuk membuka file HTML
                options = webdriver.ChromeOptions()
                options.add_argument('--headless')  # Menjalankan browser di headless mode (tanpa GUI)
                options.add_argument('--no-sandbox')
                options.add_argument('--disable-dev-shm-usage')
                options.add_argument('--force-device-scale-factor=0.65')  # Zoom level 125%


                # Inisialisasi WebDriver
                driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

                # Buka file HTML
                html_path = (os.path.join(plot, 'final.html'))
                driver.get(html_path)

                # Tunggu sebentar jika halaman memerlukan waktu untuk load
                time.sleep(2)

                # Ambil screenshot
                screenshot_path = os.path.join(plot, 'final.png')
                driver.save_screenshot(screenshot_path)

                # Tutup browser setelah selesai
                driver.quit()

                print(f'Screenshot berhasil disimpan di: {screenshot_path}')
                
                    
                    
            
            except requests.ConnectionError:
                # Do something
                print("The internet connection is down")
                x_coords = merged_csv_final['x1'] 
                y_coords= merged_csv_final['y1']
                status = merged_csv_final['Status']
                colors = {'Accurate': 'limegreen', 
                        #   'Missed': 'red', 
                            'Not Accurate': 'dodgerblue', 
                            'Not Sprayed' : 'brown'}
                # plt.figure(figsize=(10, 6))
                sns.scatterplot(x=x_coords, y=y_coords, hue=status, palette=colors, data=merged_csv_final, s=10)
                plt.xlabel('Longitude')
                plt.ylabel('Latitude')
                plt.title('Status of Tree')
                plt.grid(True)
                plt.gca().xaxis.get_major_formatter().set_useOffset(False)
                
                plt.savefig(plot_final, format='png', bbox_inches='tight', dpi=200)
                plt.close()
            self.progressBar.setValue(70)   
            # plot_final = os.path.join(plot, 'final.png')
            


            data_final = merged_csv_final['Status'].value_counts()
            # df2 = pd.DataFrame()
            # df2['Status'] = data_final.index
            # df2['Number of Trees'] = data_final.values
            # df2['Percentage'] = (df2['Number of Trees'] / df2['Number of Trees'].sum() * 100).round(2)
            # total = pd.DataFrame([['Total', df2['Number of Trees'].sum().round(0), df2['Percentage'].sum().round(0)]], columns=['Status', 'Number of Trees', 'Percentage'])
            # df2 = pd.concat([df2, total], ignore_index=True)
            

            df2 = pd.DataFrame()
            df2['Status'] = data_final.index
            df2['Number of Trees'] = data_final.values

            total_sprayed = df2[df2['Status'].isin(['Accurate', 'Not Accurate'])]['Number of Trees'].sum()
            total_mission = total_sprayed + df2[df2['Status'] == 'Not Sprayed']['Number of Trees'].sum()

            # Menghitung persentase untuk Accurate dan Not Accurate
            df2.loc[df2['Status'] == 'Accurate', 'Percentage'] = ((df2.loc[df2['Status'] == 'Accurate', 'Number of Trees'] / total_sprayed * 100).round(2).astype(str) + "%")
            df2.loc[df2['Status'] == 'Not Accurate', 'Percentage'] = ((df2.loc[df2['Status'] == 'Not Accurate', 'Number of Trees'] / total_sprayed * 100).round(2).astype(str) + "%")
            df2.loc[df2['Status'] == 'Not Sprayed', 'Percentage'] = ""

            df2 = pd.concat([df2, pd.DataFrame([
                ['Total Sprayed', total_sprayed, '100%'],
                ['Total Mission', total_mission, '']
            ], columns=['Status', 'Number of Trees', 'Percentage'])], ignore_index=True)

            # Mengatur urutan
            status_order = ["Accurate", "Not Accurate", "Not Sprayed", "Total Sprayed", "Total Mission"]
            df2['Status'] = pd.Categorical(df2['Status'], categories=status_order, ordered=True)
            df2 = df2.sort_values('Status').reset_index(drop=True)

            pdf.add_page()
            company1  = f'{self.comboBox_2.currentText()}'
            estate1 = f'{self.comboBox_1.currentText()}'
            date1 = f'{tanggal}'
            drone_id1 = f'{self.lineEdit_4a.text()}'
            accu = f'{self.comboBox_3a.currentText()}'
            pdf.summary(company1, estate1, drone_id1, date1, accu)

            pdf.add_image(plot_final, 5, 50, 190, 130)
            pdf.add_table(df2, 35, 200, 'Spray Accuration Summary')
            
            print('selesai')
            pdf.output(os.path.join(result, f"Daily Accuration Report {self.lineEdit_4a.text()} {self.comboBox_1.currentText()} {block_all} {tanggal} .pdf"))


            self.progressBar.setValue(80)
            # Membuat workbook dan worksheet_df
            wb = Workbook()
            sheet_df = wb.active
            # Rename the sheet
            sheet_df.title = 'DF'
            # Menambahkan sheet baru
            sheet_log = wb.create_sheet(title='Logs')
            # Menambahkan sheet baru
            sheet_flight_statistic = wb.create_sheet(title='Flight Statistic')

            # Mengatur lebar kolom
            sheet_df.column_dimensions['A'].width = 4
            sheet_df.column_dimensions['B'].width = 23
            sheet_df.column_dimensions['C'].width = 23
            sheet_df.column_dimensions['D'].width = 8
            sheet_df.column_dimensions['E'].width = 23
            sheet_df.column_dimensions['F'].width = 23
            sheet_df.column_dimensions['G'].width = 23
            sheet_df.column_dimensions['H'].width = 23
            sheet_df.row_dimensions[2].height = 25
            sheet_df.row_dimensions[28].height = 30
            sheet_df.row_dimensions[32].height = 60
            sheet_df.row_dimensions[33].height = 60
            sheet_df.row_dimensions[34].height = 60
            sheet_df.row_dimensions[35].height = 60
            sheet_df.row_dimensions[36].height = 60
            sheet_df.row_dimensions[37].height = 60
            sheet_df.row_dimensions[38].height = 60
            sheet_df.row_dimensions[40].height = 60
            sheet_df.row_dimensions[41].height = 60
            sheet_df.row_dimensions[42].height = 60
            sheet_df.row_dimensions[43].height = 60
            sheet_df.row_dimensions[44].height = 60
            sheet_df.row_dimensions[45].height = 60
            sheet_df.row_dimensions[46].height = 60
            sheet_df.row_dimensions[47].height = 60
            sheet_df.row_dimensions[48].height = 60
            sheet_df.row_dimensions[49].height = 60
            sheet_df.row_dimensions[50].height = 60
            sheet_df.row_dimensions[51].height = 60

            # Menggabungkan sel B2:H2 dan menambahkan teks
            sheet_df.merge_cells('B2:H2')
            sheet_df['B2'] = 'DAILY FLIGHT REPORT'
            sheet_df['B2'].alignment = Alignment(horizontal='center', vertical='center')
            sheet_df['B2'].font = Font(name='Times New Roman', size=20, bold=True)

            merge_ranges = [            
                ('C23:G23', 'C23', f'Officer: {self.lineEdit_10z3.text()}'),
                ('C24:G24', 'C24', 'Mission:'),
                ('C25:G26', 'C25', ''),
                ('C27:G27', 'C27', 'Purpose:'),
                ('C28:G29', 'C28', ''),
                ('C30:G30', 'C30', 'Outcome:'),
                ('D32:E32', 'D32', ''),
                ('F32:G32', 'F32', ''),
                ('D33:E33', 'D33', ''),
                ('F33:G33', 'F33', ''),
                ('D34:E34', 'D34', ''),
                ('F34:G34', 'F34', ''),
                ('D35:E35', 'D35', ''),
                ('F35:G35', 'F35', ''),
                ('D36:E36', 'D36', ''),
                ('F36:G36', 'F36', ''),
                ('D37:E37', 'D37', ''),
                ('F37:G37', 'F37', ''),
                ('D38:E38', 'C38', ''),
                ('F38:G38', 'F38', ''),
                ('C39:G39', 'C39', 'Working Time:'),
                ('C40:E40', 'C40', ''),
                ('F40:G40', 'F40', ''),
                ('C41:E41', 'C41', ''),
                ('F41:G41', 'F41', ''),
                ('C42:E42', 'C42', ''),
                ('F42:G42', 'F42', ''),
                ('C43:E43', 'C43', ''),
                ('F43:G43', 'F43', ''),
                ('C44:E44', 'C44', ''),
                ('F44:G44', 'F44', ''),
                ('C45:E45', 'C45', ''),
                ('F45:G45', 'F45', ''),
                ('C46:E46', 'C46', ''),
                ('F46:G46', 'F46', ''),
                ('C47:E47', 'C47', ''),
                ('F47:G47', 'F47', ''),
                ('C48:E48', 'C48', ''),
                ('F48:G48', 'F48', ''),
                ('C49:E49', 'C49', ''),
                ('F49:G49', 'F49', ''),
                ('C50:E50', 'C50', ''),
                ('F50:G50', 'F50', ''),
                ('C51:E51', 'C51', ''),
                ('F51:G51', 'F51', '')
            ]
            # Loop through merge ranges to set merged cells, texts, and fonts
            for merge_range, cell, text in merge_ranges:
                sheet_df.merge_cells(merge_range)
                sheet_df[cell] = text
                sheet_df[cell].font = Font(name='Calibri', size=11, bold=True)

            thin_border = Border(left=Side(style='thin'), 
                                right=Side(style='thin'), 
                                top=Side(style='thin'), 
                                bottom=Side(style='thin'))

            highlight_ranges = [
                (5, 10, 3, 3),  # C5:D5
                (12, 15, 3, 3),   # G5:H5
                (17, 19, 3, 3),   # G5:H5
                
                (5, 10, 6, 6),   # G5:H5
                (12, 16, 6, 6),   # G5:H5
                (18, 19, 6, 6),   # G5:H5
                
                (23, 23, 3, 7),   # G5:H5
                (30, 30, 3, 7),   # G5:H5
                (31, 31, 3, 7),   # G5:H5   
                (31, 51, 3, 7),   # G5:H5
                ]
            fill = PatternFill(start_color='DEEAF6', end_color='DEEAF6', fill_type='solid')
            for min_row, max_row, min_col, max_col in highlight_ranges:
                for row in sheet_df.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                    for cell in row:
                        cell.border = thin_border
                        cell.fill = fill

            # Define the border style
            thin_border = Side(style='thin')
            start_row = 24
            end_row = 29
            start_col = 3  # Column C
            end_col = 7   # Column H
            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    cell = sheet_df.cell(row=row, column=col)
                    border = Border(
                        top=thin_border if row == start_row else None,
                        bottom=thin_border if row == end_row else None,
                        left=thin_border if col == start_col else None,
                        right=thin_border if col == end_col else None)
                    cell.border = border
                    cell.fill = fill

            # Data yang akan diisi
            data = [
                ('B5', 'Arrival Time to Location'),
                ('B6', 'Height From Crop (m)'),
                ('B7', 'Flying Speed (m/s)'),
                ('B8', 'Swath (m)'),
                ('B9', 'Customer'),
                ('B10', 'Project'),
                        
                ('B12', 'Drone Type'),
                ('B13', 'Herbicide Ratio (%)'),
                ('B14', 'Chemical Type'),
                ('B15', 'Chemical Mixture'),
                
                ('B17', 'Driver'),
                ('B18', 'Car Plate No'),
                ('B19', 'Team'),

                ('E5', 'Herbicide Time Ready'),
                ('E6', 'PIC'),
                ('E7', 'Pilot'),
                ('E8', 'Co-Pilot/Coordinator'),
                ('E9', 'GCS'),
                ('E10', 'Support'),
            
                ('E12', 'Nozzle Type (Color)'),
                ('E13', 'District Name'),
                ('E14', 'District PIC'),
                ('E15', 'Spraying'),
                ('E16', 'Rotation num'),
                
                ('E18', 'Date'),
                ('E19', 'Contour Degree'),
                
                ('B23', 'Note/Info'),
                ('C31', 'No'), ('D31', 'Situation'), ('F31', 'Solution')
                ]
            # Font yang akan digunakan
            font = Font(name='Calibri', size=11, bold=True)
            for cell, value in data:
                sheet_df[cell] = value
                sheet_df[cell].font = font
            data = [
                ('C32', '1'),
                ('C33', '2'),
                ('C34', '3'),
                ('C35', '4'),
                ('C36', '5'),
                ('C37', '6'),
                ('C38', '7'),]
            font = Font(name='Calibri', size=11)
            for cell, value in data:
                sheet_df[cell] = value
                sheet_df[cell].font = font

            # Menyelaraskan teks ke tengah (horizontal dan vertikal)
            alignment = Alignment(horizontal='center', vertical='center')
            sheet_df['C31'].alignment = alignment
            sheet_df['D31'].alignment = alignment
            sheet_df['F31'].alignment = alignment
            sheet_df['C32'].alignment = alignment
            sheet_df['C33'].alignment = alignment
            sheet_df['C34'].alignment = alignment
            sheet_df['C35'].alignment = alignment
            sheet_df['C36'].alignment = alignment
            sheet_df['C36'].alignment = alignment
            sheet_df['C37'].alignment = alignment
            sheet_df['C38'].alignment = alignment
            sheet_df.merge_cells('D31:E31')
            sheet_df.merge_cells('F31:G31')

            # sheet_log
            sheet_log.column_dimensions['A'].width = 13
            sheet_log.column_dimensions['B'].width = 16
            sheet_log.column_dimensions['C'].width = 24
            sheet_log.column_dimensions['D'].width = 13
            sheet_log.column_dimensions['E'].width = 13
            sheet_log.column_dimensions['F'].width = 13
            sheet_log.column_dimensions['G'].width = 13
            sheet_log.column_dimensions['H'].width = 26
            sheet_log.row_dimensions[1].height = 35

            data_judul_sheet_log = [
                ('A1', 'Username'),
                ('B1', 'Start Time'),
                ('C1', 'Block Area Name'),
                ('D1', 'Flight Time'),
                ('E1', 'Spray Trees'),
                ('F1', 'Spray Area Ha'),
                ('G1', 'Spray Dose (L)'),
                ('H1', 'Drone Id')
            ]
            font = Font(name='Arial', size=11, bold=True)
            for cell, value in data_judul_sheet_log:
                sheet_log[cell] = value
                sheet_log[cell].font = font

            sheet_log['A1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            sheet_log['B1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            sheet_log['C1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            sheet_log['D1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            sheet_log['E1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            sheet_log['F1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            sheet_log['G1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            sheet_log['H1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            sheet_log['I1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

            fill_2 = PatternFill(start_color='FDE9D9', end_color='FDE9D9', fill_type='solid')
            fill_3 = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

            highlight_ranges_sheet_log = [(1, 1, 1, 8)]
            for min_row, max_row, min_col, max_col in highlight_ranges_sheet_log:
                for row in sheet_log.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                    for cell in row:
                        # cell.border = thin_border
                        cell.fill = fill_2

            highlight_ranges_sheet_log = [(132, 132, 4, 7), (133,133, 4, 5)]

            for min_row, max_row, min_col, max_col in highlight_ranges_sheet_log:
                for row in sheet_log.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                    for cell in row:
                        # cell.border = thin_border
                        cell.fill = fill_3

            sheet_log.merge_cells('A132:C133')
            sheet_log['A132'] = 'Total'
            sheet_log['A132'].alignment = Alignment(horizontal='center', vertical='center')
            sheet_log['A132'].font = Font(name='Arial', size=11, bold=True)

            sheet_log.merge_cells('D133:E133')
            sheet_log['D133'] = 'Flight'
            sheet_log['D133'].alignment = Alignment(horizontal='center', vertical='center')
            sheet_log['D133'].font = Font(name='Arial', size=11, bold=True)
            
            
            sheet_log.merge_cells('D138:G138')
            sheet_log['D138'] = '\u2193 Leave it blank if there is only one plot'
            sheet_log['D138'].font = Font(name='Arial', size=11)

            sheet_log['C138'] = "\u2193 Fill DF's Plot Numbers"
            sheet_log['C138'].font = Font(name='Arial', size=11)

            fill_4 = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
            highlight_ranges_sheet_log = [(139, 139, 3, 7)]
            for min_row, max_row, min_col, max_col in highlight_ranges_sheet_log:
                for row in sheet_log.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                    for cell in row:
                        # cell.border = thin_border
                        cell.fill = fill_4

            fill_5 = PatternFill(start_color='B8CCE4', end_color='B8CCE4', fill_type='solid')
            highlight_ranges_sheet_log = [(140, 144, 3, 7)]
            for min_row, max_row, min_col, max_col in highlight_ranges_sheet_log:
                for row in sheet_log.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                    for cell in row:
                        # cell.border = thin_border
                        cell.fill = fill_5

            sheet_log['C139'] = "Plot name"
            sheet_log['C139'].font = Font(name='Arial', size=11, color="FFFFFF")
            sheet_log['D139'] = "Flight time"
            sheet_log['D139'].font = Font(name='Arial', size=11, color="FFFFFF")
            sheet_log['E139'] = "Spray trees"
            sheet_log['E139'].font = Font(name='Arial', size=11, color="FFFFFF")
            sheet_log['F139'] = "Spray Area"
            sheet_log['F139'].font = Font(name='Arial', size=11, color="FFFFFF")
            sheet_log['G139'] = "Spray Dose"
            sheet_log['G139'].font = Font(name='Arial', size=11, color="FFFFFF")

            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            thin_border_2 = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))

            highlight_ranges = [(1, 131, 1, 8), (140, 144, 3, 7)]
            for min_row, max_row, min_col, max_col in highlight_ranges:
                for row in sheet_log.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                    for cell in row:
                        cell.border = thin_border

            highlight_ranges = [(132, 133, 1, 5), (132, 132, 3, 7)]
            for min_row, max_row, min_col, max_col in highlight_ranges:
                for row in sheet_log.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                    for cell in row:
                        cell.border = thin_border_2

            # thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            # thin_border_2 = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))

            # highlight_ranges = [(1, 131, 1, 8)]
            # for min_row, max_row, min_col, max_col in highlight_ranges:
            #     for row in sheet_log.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            #         for cell in row:
            #             cell.border = thin_border

            # highlight_ranges = [(132, 133, 1, 5), (132, 132, 3, 7)]
            # for min_row, max_row, min_col, max_col in highlight_ranges:
            #     for row in sheet_log.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            #         for cell in row:
            #             cell.border = thin_border_2



            sheet_flight_statistic.merge_cells('A2:J2')
            sheet_flight_statistic['A2'] = 'FLIGHT STATISTIC'
            sheet_flight_statistic['A2'].alignment = Alignment(horizontal='center', vertical='center')
            sheet_flight_statistic['A2'].font = Font(name='Times New Roman', size=20, bold=True)

            sheet_flight_statistic.merge_cells('G10:H10')
            sheet_flight_statistic.merge_cells('G11:H11')
            sheet_flight_statistic.merge_cells('G12:H12')
            sheet_flight_statistic.merge_cells('G13:H13')

            # Data yang akan diisi
            data = [
                    ('G10', 'Unit'),
                    ('G11', 'Working Area'),
                    ('G12', 'Sprayed Area'),
                    ('G13', 'Flight Time'),
                    ('G14', 'Total Chemical'),
                    # ('J11', 'Ha'),
                    ('J12', 'Ha'),
                    ('J13', 'Hours'),
                    ('J14', 'Liter')
                    ]

            # Font yang akan digunakan
            center_alignment = Alignment(horizontal='center', vertical='center')
            font = Font(name='Arial Narrow', size=11, bold=True)
            for cell, value in data:
                sheet_flight_statistic[cell] = value
                sheet_flight_statistic[cell].font = font
                sheet_flight_statistic[cell].alignment = center_alignment

            # Define the border style
            center_alignment_left = Alignment(horizontal='left', vertical='center')
            thin_border = Side(style='medium')
            start_row = 7
            end_row = 14
            start_col = 7  # Column C
            end_col = 10   # Column H
            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    cell = sheet_flight_statistic.cell(row=row, column=col)
                    border = Border(
                        top=thin_border if row == start_row else None,
                        bottom=thin_border if row == end_row else None,
                        left=thin_border if col == start_col else None,
                        right=thin_border if col == end_col else None)
                    cell.border = border
                    cell.alignment = center_alignment_left
                    # cell.fill = fill

            sheet_flight_statistic.merge_cells('G7:J7')
            sheet_flight_statistic['G7'] = self.comboBox_1.currentText()
            sheet_flight_statistic['G7'].alignment = Alignment(horizontal='center', vertical='center')
            sheet_flight_statistic['G7'].font = Font(name='Arial', size=11, bold=True)

            sheet_flight_statistic.merge_cells('I8:J8')
            sheet_flight_statistic['I8'].value = df_merged['tanggal'].head().iloc[0]
            sheet_flight_statistic['I8'].alignment = Alignment(horizontal='center', vertical='center')
            sheet_flight_statistic['I8'].font = Font(name='Arial', size=11, bold=True)

            sheet_flight_statistic.merge_cells('G8:H8')
            sheet_flight_statistic['G8'] = 'Work Date'
            sheet_flight_statistic['G8'].alignment = Alignment(horizontal='center', vertical='center')
            sheet_flight_statistic['G8'].font = Font(name='Arial', size=11, bold=True)

            sheet_flight_statistic.merge_cells('G9:J9')
            sheet_flight_statistic['G9'] = self.comboBox_2.currentText()
            sheet_flight_statistic['G9'].alignment = Alignment(horizontal='center', vertical='center')
            sheet_flight_statistic['G9'].font = Font(name='Arial', size=11, bold=True)



            #data df sheet
            # df_merged['tanggal'] = pd.to_datetime(df_merged['tanggal']).dt.strftime('%d-%B-%Y')
            df_merged['tanggal'] = pd.to_datetime(df_merged['tanggal'], format='%d-%B-%Y')
            formatted_datee = df_merged['tanggal'].mode()[0].strftime('%Y/%m/%d')
            
            data_df_atas = [
                ('C5' , self.lineEdit_10a.text()),
                ('C6' , self.lineEdit_10b.text()),
                ('C7' , self.lineEdit_10c.text()),
                ('C8' , self.lineEdit_10d.text()),
                ('C9' , self.comboBox_2.currentText()),
                ('C10', self.lineEdit_combo6.text()),
                
                ('C12', self.lineEdit_4a.text()),
                ('C13', self.lineEdit_10f.text()),
                ('C14', self.comboBox_10g.currentText()),
                ('C15', self.lineEdit_10h.text()),
                ('C17', self.lineEdit_10i.text()),
                ('C18', self.lineEdit_10j.text()),
                ('C19', self.lineEdit_team.text()),
                
                ('F5' , self.lineEdit_10k.text()),       
                ('F6' , self.lineEdit_10l.text()),
                ('F7' , self.lineEdit_10m.text()),
                ('F8' , self.lineEdit_10n.text()),
                ('F9' , self.lineEdit_10o.text()),
                ('F10', self.lineEdit_10p.text()),
                
                ('F12', self.lineEdit_10q.text()),
                ('F13', self.lineEdit_10r.text()),
                ('F14', self.lineEdit_10s.text()),
                ('F15', self.spinBox_10t.value()),
                ('F16', self.comboBox_10t.currentText()),
                
                ('F18', formatted_datee),
                ('F19', self.comboBox_10z1.currentText()),
                
                ('C25', self.lineEdit_10z4.text()),                 #Isinya Mision
                ('C28', (f"To complement the spraying with the ‘palm by palm’ method on plot number {block_all}")),          #Isinya Purpose 
                
                ('D32', self.textEdit_situation_1.toPlainText()),
                ('D33', self.textEdit_situation_2.toPlainText()),
                ('D34', self.textEdit_situation_3.toPlainText()),
                ('D35', self.textEdit_situation_4.toPlainText()),
                ('F32', self.textEdit_solution_1.toPlainText()),
                ('F33', self.textEdit_solution_2.toPlainText()),
                ('F34', self.textEdit_solution_3.toPlainText()),
                ('F35', self.textEdit_solution_4.toPlainText()),
                ('C40', self.textEdit_worktime_1a.toPlainText()),
                ('F40', self.textEdit_worktime_1b.toPlainText()),
                ('C41', self.textEdit_worktime_2a.toPlainText()),
                ('F41', self.textEdit_worktime_2b.toPlainText()),
                ('C42', self.textEdit_worktime_3a.toPlainText()),
                ('F42', self.textEdit_worktime_3b.toPlainText()),
                ('C43', self.textEdit_worktime_4a.toPlainText()),
                ('F43', self.textEdit_worktime_4b.toPlainText()),
                ('C44', self.textEdit_worktime_5a.toPlainText()),
                ('F44', self.textEdit_worktime_5b.toPlainText()),
                ('C45', self.textEdit_worktime_6a.toPlainText()),
                ('F45', self.textEdit_worktime_6b.toPlainText()),
                ('C46', self.textEdit_worktime_7a.toPlainText()),
                ('F46', self.textEdit_worktime_7b.toPlainText())]
            
            
            alignment_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            font = Font(name='Calibri', size=11)
            def set_value(sheet, cell, value):
                if isinstance(sheet[cell], openpyxl.cell.cell.Cell):  # Pastikan sel adalah sel utama
                    sheet[cell] = value 
                    sheet[cell].font = font
                    sheet[cell].alignment = alignment_center
            for cell, value in data_df_atas:
                set_value(sheet_df, cell, value)

            data_isian_df = [
                ('D32', self.textEdit_situation_1.toPlainText()),
                ('G32', self.textEdit_solution_1.toPlainText()),
                ('D33', self.textEdit_situation_2.toPlainText()),
                ('G33', self.textEdit_solution_2.toPlainText()),
                ('D34', self.textEdit_situation_3.toPlainText()),
                ('G34', self.textEdit_solution_3.toPlainText()),
                ('D35', self.textEdit_situation_4.toPlainText()),
                ('G35', self.textEdit_solution_4.toPlainText()),]
            alignment_kiri = Alignment(horizontal='left', vertical='center', wrap_text=True)
            font = Font(name='Calibri', size=11)
            def set_value(sheet, cell, value):
                if isinstance(sheet[cell], openpyxl.cell.cell.Cell):  # Pastikan sel adalah sel utama
                    sheet[cell] = value
                    sheet[cell].font = font
                    sheet[cell].alignment = alignment_kiri
            for cell, value in data_isian_df:
                set_value(sheet_df, cell, value)


            data_df_atas = [
                # ('A2:A25' , self.lineEdit_10a.text()),
                ('C6' , self.lineEdit_10b.text()),
                ('C7' , self.lineEdit_10c.text())]
            alignment_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            font = Font(name='Calibri', size=11)
            def set_value(sheet, cell, value):
                if isinstance(sheet[cell], openpyxl.cell.cell.Cell):  # Pastikan sel adalah sel utama
                    sheet[cell] = value 
                    sheet[cell].font = font
                    sheet[cell].alignment = alignment_center
            for cell, value in data_df_atas:
                set_value(sheet_df, cell, value)

            
            #logs sheet
            self.progressBar.setValue(90)
            awal = pd.read_excel(os.path.join(junk, 'GABUNG.xlsx'))

            source_column = 'nama'
            extracted_data_log = []

            sph_rumus = {blok1 : sph1, blok2: sph2, blok3: sph3, blok4: sph4, blok5:sph5}
            dose_rumus = {blok1: dosis1, blok2: dosis2, blok3: dosis3, blok4: dosis4, blok5: dosis5}
            
            
            for index, row in awal.iterrows():
                block_area_name = row[source_column]
                match_found = False  # Flag untuk memeriksa kecocokan

                for text in sph_rumus.keys():
                    if isinstance(block_area_name, str) and text in block_area_name:
                        # Menghitung nilai SPH dan Dose berdasarkan rumus
                        sph_value = row['total'] / sph_rumus[text]
                        dose_value = row['total'] * (dose_rumus[text] / 1000)

                        # Membuat dictionary untuk menyimpan data baris
                        row_data = {
                            'nama': block_area_name,
                            'total': row['total'],
                            'start': row['start'],
                            'selisih': row['selisih'],
                            'tanggal': row['tanggal'],
                            'SPH': sph_value,
                            'dose': dose_value
                        }
                        extracted_data_log.append(row_data)
                        match_found = True  # Menandai kecocokan ditemukan
                        break  # Keluar dari loop jika teks sudah cocok

                # Menangani data di luar blok
                if not match_found:
                    row_data = {
                        'nama': block_area_name,
                        'total': row['total'],
                        'start': row['start'],
                        'selisih': row['selisih'],
                        'tanggal': row['tanggal'],
                        'SPH': None,  # Nilai SPH tidak dihitung
                        'dose': None  # Nilai dose tidak dihitung
                    }
                    extracted_data_log.append(row_data)


            awal1= pd.DataFrame(extracted_data_log)
            
            
            start_misi = awal1 ['start']
            durasi_misi = awal1['selisih']
            nama_misi = awal1['nama']
            pohon_jumlah = awal1['total']
            sph_tampil = awal1['SPH']
            dosis_tampil = awal1['dose']
            
            
            for row in range(2, 131):
                sheet_log[f'A{row}'] = self.lineEdit_4a.text().upper()

            start_row = 2   
            end_row = 131
            for i in range(len(start_misi)):
                sheet_log[f'B{start_row + i}'] = start_misi[i]
                sheet_log[f'C{start_row + i}'] = nama_misi[i]
                sheet_log[f'D{start_row + i}'] = durasi_misi[i]
                sheet_log[f'E{start_row + i}'] = pohon_jumlah[i]
                sheet_log[f'F{start_row + i}'] = sph_tampil[i]                
                sheet_log[f'G{start_row + i}'] = dosis_tampil[i]                
                    
                    
            source_column = 'C'
            extracted_data = []
            extracted_data1 = []
            extracted_data2 = []
            extracted_data3 = []
            extracted_data4 = []
            
            for row1 in range(start_row, end_row):
                block_area_name = sheet_log[f'{source_column}{row1}'].value

                # Check if the block_area_name is not None and contains 'P23B'
                if isinstance(block_area_name, str) and blok1 in block_area_name:
                # if block_area_name and text_10v in block_area_name:
                    # Create a dictionary to hold the row data
                    row_data = {
                        'Plot name': block_area_name,
                        'Flight time': sheet_log[f'D{row1}'].value,
                        'Spray trees': sheet_log[f'E{row1}'].value,
                        'Spray Area': sheet_log[f'F{row1}'].value,
                        'Spray Dose': sheet_log[f'G{row1}'].value
                    }
                    # Append the row data to the extracted_data list
                    extracted_data2.append(row_data)
                    
                
                    
            # log5 = pd.DataFrame(extracted_data2) #second plot
            # print(log5)     
            
            # if not log5.empty:
            #     durasi_misi1 = log5['Flight time']
            #     def parse_duration(time_str):
            #         # Mengonversi string waktu ke objek timedelta
            #         t = datetime.datetime.strptime(time_str, "%H:%M:%S")
            #         return timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
            #     # Inisialisasi total durasi sebagai timedelta nol
            #     total_duration1 = timedelta()
            #     # Menambahkan setiap durasi ke total_duration
            #     for duration1 in durasi_misi1:
            #         total_duration1 += parse_duration(duration1)


            #     sheet_log['E140'] = log5['Spray trees'].sum()
            #     sheet_log['F140'] = log5['Spray Area'].sum()
            #     sheet_log['D140'] = total_duration1
            #     sheet_log['G140'] = log5['Spray Dose'].sum()
            #     sheet_log['C140'] = blok1
            #     pass
            # else:
            #     print("DataFrame df kosong. Tidak ada data untuk P23B.")
            #     pass
            
            log5 = pd.DataFrame(extracted_data2)  # Second plot
            print(log5)


            if not log5.empty:
                durasi_misi1 = log5['Flight time'].dropna()  # Abaikan NaN dalam kolom 'Flight time'

                def parse_duration(time_str):
                    if not isinstance(time_str, str) or not time_str.strip():
                        raise ValueError(f"Invalid time format: {time_str}")
                    t = datetime.strptime(time_str, "%H:%M:%S")
                    return timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
                
                # def parse_duration(time_str):
                #     # Mengonversi string waktu ke objek timedelta
                #     t = datetime.strptime(time_str, "%H:%M:%S")
                #     return timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)

                # Inisialisasi total durasi sebagai timedelta nol
                total_duration1 = timedelta()

                # Menambahkan setiap durasi yang valid ke total_duration
                for duration1 in durasi_misi1:
                    try:
                        total_duration1 += parse_duration(duration1)
                    except ValueError as e:
                        print(f"Skipping invalid time: {e}")

                    print(f"Total duration: {total_duration1}")
                    # total_duration1 += parse_duration(duration1)

                # Menyimpan hasil ke sheet_log
                sheet_log['E140'] = log5['Spray trees'].sum()
                sheet_log['F140'] = log5['Spray Area'].sum()
                sheet_log['D140'] = total_duration1
                sheet_log['G140'] = log5['Spray Dose'].sum()
                sheet_log['C140'] = blok1
            else:
                print("DataFrame log5 kosong. Tidak ada data untuk P23B.")

            
            
            
            for row2 in range(start_row, end_row):
                block_area_name = sheet_log[f'{source_column}{row2}'].value

                # Check if the block_area_name is not None and contains 'P23B'
                if isinstance(block_area_name, str) and blok2 in block_area_name:
                # if block_area_name and text_10v in block_area_name:
                    # Create a dictionary to hold the row data
                    row_data = {
                        'Plot name': block_area_name,
                        'Flight time': sheet_log[f'D{row2}'].value,
                        'Spray trees': sheet_log[f'E{row2}'].value,
                        'Spray Area': sheet_log[f'F{row2}'].value,
                        'Spray Dose': sheet_log[f'G{row2}'].value
                    }
                    # Append the row data to the extracted_data list
                    extracted_data.append(row_data)
                    
                
                    
            log2 = pd.DataFrame(extracted_data) #second plot
            print(log2)     
            
            if not log2.empty:
                durasi_misi1 = log2['Flight time'].dropna()
                def parse_duration(time_str):
                    # Mengonversi string waktu ke objek timedelta
                    t = datetime.strptime(time_str, "%H:%M:%S")
                    return timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
                # Inisialisasi total durasi sebagai timedelta nol
                total_duration1 = timedelta()
                # Menambahkan setiap durasi ke total_duration
                for duration1 in durasi_misi1:
                    total_duration1 += parse_duration(duration1)


                sheet_log['E141'] = log2['Spray trees'].sum()
                sheet_log['F141'] = log2['Spray Area'].sum()
                sheet_log['D141'] = total_duration1
                sheet_log['G141'] = log2['Spray Dose'].sum()
                sheet_log['C141'] = blok2
                pass
            else:
                print("DataFrame df kosong. Tidak ada data untuk P23B.")
                pass
            
            
            
            
            for row3 in range(start_row, end_row):
                block_area_name = sheet_log[f'{source_column}{row3}'].value

                # Check if the block_area_name is not None and contains 'P23B'
                if isinstance(block_area_name, str) and blok3 in block_area_name:
                # if block_area_name and text_10w in block_area_name:
                    # Create a dictionary to hold the row data
                    row_data = {
                        'Plot name': block_area_name,
                        'Flight time': sheet_log[f'D{row3}'].value,
                        'Spray trees': sheet_log[f'E{row3}'].value,
                        'Spray Area': sheet_log[f'F{row3}'].value,
                        'Spray Dose': sheet_log[f'G{row3}'].value
                    }
                    # Append the row data to the extracted_data list
                    extracted_data1.append(row_data)
                    
            log3 = pd.DataFrame(extracted_data1) #second plot
            print(log3)     
            
            if not log3.empty:
                durasi_misi2 = log3['Flight time'].dropna()
                def parse_duration(time_str):
                    # Mengonversi string waktu ke objek timedelta
                    t = datetime.strptime(time_str, "%H:%M:%S")
                    return timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
                # Inisialisasi total durasi sebagai timedelta nol
                total_duration2 = timedelta()
                # Menambahkan setiap durasi ke total_duration
                for duration2 in durasi_misi2:
                    total_duration2 += parse_duration(duration2)


                sheet_log['E142'] = log3['Spray trees'].sum()
                sheet_log['F142'] = log3['Spray Area'].sum()
                sheet_log['D142'] = total_duration2
                sheet_log['G142'] = log3['Spray Dose'].sum()
                sheet_log['C142'] = blok3
                pass
            else:
                print("DataFrame df kosong. Tidak ada data untuk P23B.")
                pass
            




            for row4 in range(start_row, end_row):
                block_area_name = sheet_log[f'{source_column}{row4}'].value

                # Check if the block_area_name is not None and contains 'P23B'
                if isinstance(block_area_name, str) and blok4 in block_area_name:
                # if block_area_name and text_10w in block_area_name:
                    # Create a dictionary to hold the row data
                    row_data = {
                        'Plot name': block_area_name,
                        'Flight time': sheet_log[f'D{row4}'].value,
                        'Spray trees': sheet_log[f'E{row4}'].value,
                        'Spray Area': sheet_log[f'F{row4}'].value,
                        'Spray Dose': sheet_log[f'G{row4}'].value
                    }
                    # Append the row data to the extracted_data list
                    extracted_data3.append(row_data)
                    
            log6 = pd.DataFrame(extracted_data3) #second plot
            print(log6)     
            
            if not log6.empty:
                durasi_misi2 = log6['Flight time'].dropna()
                def parse_duration(time_str):
                    # Mengonversi string waktu ke objek timedelta
                    t = datetime.strptime(time_str, "%H:%M:%S")
                    return timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
                # Inisialisasi total durasi sebagai timedelta nol
                total_duration2 = timedelta()
                # Menambahkan setiap durasi ke total_duration
                for duration2 in durasi_misi2:
                    total_duration2 += parse_duration(duration2)


                sheet_log['E143'] = log6['Spray trees'].sum()
                sheet_log['F143'] = log6['Spray Area'].sum()
                sheet_log['D143'] = total_duration2
                sheet_log['G143'] = log6['Spray Dose'].sum()
                sheet_log['C143'] = blok4
                pass
            else:
                print("DataFrame df kosong. Tidak ada data untuk P23B.")
                pass
            
            
            
            for row3 in range(start_row, end_row):
                block_area_name = sheet_log[f'{source_column}{row3}'].value

                # Check if the block_area_name is not None and contains 'P23B'
                if isinstance(block_area_name, str) and blok5 in block_area_name:
                # if block_area_name and text_10w in block_area_name:
                    # Create a dictionary to hold the row data
                    row_data = {
                        'Plot name': block_area_name,
                        'Flight time': sheet_log[f'D{row3}'].value,
                        'Spray trees': sheet_log[f'E{row3}'].value,
                        'Spray Area': sheet_log[f'F{row3}'].value,
                        'Spray Dose': sheet_log[f'G{row3}'].value
                    }
                    # Append the row data to the extracted_data list
                    extracted_data4.append(row_data)
                    
            log7 = pd.DataFrame(extracted_data4) #second plot
            print(log7)     
            
            if not log7.empty:
                durasi_misi2 = log7['Flight time'].dropna()
                def parse_duration(time_str):
                    # Mengonversi string waktu ke objek timedelta
                    t = datetime.strptime(time_str, "%H:%M:%S")
                    return timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
                # Inisialisasi total durasi sebagai timedelta nol
                total_duration2 = timedelta()
                # Menambahkan setiap durasi ke total_duration
                for duration2 in durasi_misi2:
                    total_duration2 += parse_duration(duration2)


                sheet_log['E144'] = log7['Spray trees'].sum()
                sheet_log['F144'] = log7['Spray Area'].sum()
                sheet_log['D144'] = total_duration2
                sheet_log['G144'] = log7['Spray Dose'].sum()
                sheet_log['C144'] = blok5
                pass
            else:
                print("DataFrame df kosong. Tidak ada data untuk P23B.")
                pass
            
            
            
            


            # def parse_duration(time_str):
            #     if pd.isna(time_str):
            #         # return datetime.timedelta(0)  # Durasi nol untuk nilai kosong
            #         return timedelta(0)

            #     # Ubah time_str ke string jika bukan string
            #     if not isinstance(time_str, str):
            #         time_str = str(time_str)

            #     # Coba konversi format jam jika sudah dalam bentuk string
            #     try:
            #         # t = datetime.datetime.strptime(time_str, "%H:%M:%S")
            #         t = datetime.strptime(time_str, "%H:%M:%S")
            #         # return datetime.timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
            #         return timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
            #     except ValueError:
            #         print(f"Invalid time format: {time_str}")
            #         # return datetime.timedelta(0)  # Kembalikan durasi nol jika format salah
            #         return timedelta(0)  # Kembalikan durasi nol jika format salah
                
        
        
            # # Inisialisasi total durasi sebagai timedelta nol
            # total_duration_all = timedelta()
            
            # # Iterate through rows 2 to 30 (remember that openpyxl uses 1-based indexing)
            # for row in range(2, 131):
            #     cell_value = sheet_log[f'D{row}'].value  # Column D corresponds to 'Flight Time'
                
            #     if cell_value:  # Check if the cell is not empty
            #         total_duration_all += parse_duration(cell_value)


            # alignment_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            # font = Font(name='Arial', size=11)
            # def set_value(sheet, cell, value):
            #     if isinstance(sheet[cell], openpyxl.cell.cell.Cell):  # Pastikan sel adalah sel utama
            #         sheet[cell] = value 
            #         sheet[cell].font = font
            #         sheet[cell].alignment = alignment_center
                    
            # data_df_atas = [
            #     ('D132' , total_duration_all)
            # ]

            # for cell, value in data_df_atas:
            #     set_value(sheet_log, cell, value)
            
            
            for row in sheet_log.iter_rows(min_row=2, max_row=132, min_col=1, max_col=7):
                for cell in row:
                    cell.font = font
                    cell.alignment = alignment_center
                    
            sheet_log['A132'].alignment = Alignment(horizontal='center', vertical='center')
            sheet_log['A132'].font = Font(name='Arial', size=11, bold=True)


            #flight sheet
            for row in range(1, 16):  # Baris 1 hingga 15
                sheet_flight_statistic.row_dimensions[row].height = 27

            # Atur lebar kolom untuk setiap kolom dari A sampai dengan J
            for col in range(1, 11):  # Kolom A (1) hingga J (10)
                col_letter = openpyxl.utils.get_column_letter(col)
                sheet_flight_statistic.column_dimensions[col_letter].width = 19

            
            image_path = plot_final
            img = Image(image_path)
            img.width = 750
            img.height = 600

            # Set posisi gambar ke cell A7
            img.anchor = 'A6'
            # Sisipkan gambar ke sheet
            sheet_flight_statistic.add_image(img)
        
            # Menyimpan workbook ke file
            wb.save(os.path.join(junk, f"Terra Drone-DF Form2-Drone-{self.lineEdit_4a.text()}_ Team-{self.comboBox_2.currentText()} {self.comboBox_1.currentText()}_{tanggal}.xlsx"))
            
            
            def get_numeric_value(cell):
                try:
                    return float(cell.value)
                except (TypeError, ValueError):
                    return 0
            file_path = os.path.join(junk, f"Terra Drone-DF Form2-Drone-{self.lineEdit_4a.text()}_ Team-{self.comboBox_2.currentText()} {self.comboBox_1.currentText()}_{tanggal}.xlsx")
            wb = load_workbook(file_path, data_only = True)
            sheet_log1 = wb['Logs']
            sheet_flight_statistic1 = wb['Flight Statistic']     
            
            
            def parse_duration(time_str):
                if pd.isna(time_str):
                    # return datetime.timedelta(0)  # Durasi nol untuk nilai kosong
                    return timedelta(0)

                # Ubah time_str ke string jika bukan string
                if not isinstance(time_str, str):
                    time_str = str(time_str)

                # Coba konversi format jam jika sudah dalam bentuk string
                try:
                    # t = datetime.datetime.strptime(time_str, "%H:%M:%S")
                    t = datetime.strptime(time_str, "%H:%M:%S")
                    # return datetime.timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
                    return timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
                except ValueError:
                    print(f"Invalid time format: {time_str}")
                    # return datetime.timedelta(0)  # Kembalikan durasi nol jika format salah
                    return timedelta(0)  # Kembalikan durasi nol jika format salah
                
        
        
            # Inisialisasi total durasi sebagai timedelta nol
            total_duration_all = timedelta()
        
            # Iterate through rows 2 to 30 (remember that openpyxl uses 1-based indexing)
            for row in range(2, 131):
                cell_value = sheet_log1[f'D{row}'].value  # Column D corresponds to 'Flight Time'
                
                if cell_value:  # Check if the cell is not empty
                    total_duration_all += parse_duration(cell_value)

            # Loop untuk menjumlahkan kolom Spray Trees (E), Spray Area Ha (F), dan Spray Dose (L) (G) dari baris 2 hingga 31
            spray_trees_total = 0
            spray_area_ha_total = 0
            spray_dose_total = 0

            for row in range(2, 131):  # Baris 2 sampai baris 31
                spray_trees_total += get_numeric_value(sheet_log1[f'E{row}'])
                spray_area_ha_total += get_numeric_value(sheet_log1[f'F{row}'])
                spray_dose_total += get_numeric_value(sheet_log1[f'G{row}'])

            sheet_log1['E132'] = "=SUM(E2:E131)"
            sheet_log1['F132'] = "=SUM(F2:F131)"
            sheet_log1['G132'] = "=SUM(G2:G131)"
                    
            sheet_flight_statistic1['I10'].value = sheet_df['C12'].value
            sheet_flight_statistic1['I11'].value = block_all
            sheet_flight_statistic1['I12'].value = spray_area_ha_total
            sheet_flight_statistic1['I13'].value = total_duration_all
            sheet_flight_statistic1['I14'].value = spray_dose_total


            wb.save(os.path.join(result, f"Terra Drone-DF Form2-Drone-{self.lineEdit_4a.text()}_ Team-{self.comboBox_2.currentText()} {self.comboBox_1.currentText()}_{tanggal}.xlsx"))        
            shutil.rmtree(junk)
            self.progressBar.setFormat("Completed")
            self.progressBar.setValue(100)
        
        except Exception as e:
            QMessageBox.critical(None, "Error", f"Terjadi kesalahan yang tidak diketahui: {str(e)}")



# ------------------------------------------------------------------------ end of excel ----------------------------------------------------------

    def merge_data(self):
        try :
            print(f"Lokasi folder log adalah {self.data_pdf}")
            print(f"Lokasi folder log adalah {self.data_output_pdfmerge}")
            print(f"Lokasi Estate adalah {self.lineEdit_20.text()}")
            print(f"Lokasi Petak adalah {self.comboBox_4.currentText()}")
            print(f"Ratation Petak Ke- {self.spinBox_1.value()}")
            # print(f"Tanggal Start Penyemprotan {self.DATE_2.date().toString('dd-MMMM-yyyy')}")
            # print(f"Tanggal End/Finish Penyemprotan {self.DATE_3.date().toString('dd-MMMM-yyyy')}")

            company = self.comboBox_4a.currentText()
            estate = self.comboBox_4.currentText()
            petak = self.lineEdit_20.text()
            rotasi = self.spinBox_1.value()
            # start = self.DATE_2.date().toString('dd-MMMM-yyyy')
            # end = self.DATE_3.date().toString('dd-MMMM-yyyy')
            hektaran = self.lineEdit_21.text()
            blok1 = self.lineEdit_20.text()

            # Define a class for the PDF document
            class PDF(FPDF):
                def header(self):
                    self.set_font('Arial', 'B', 14)
                    self.cell(0, 5, 'SUMMARY ACCURATION STATUS REPORT', 0, 1, 'C')
                    
                    self.set_font('Arial', '', 12)
                    self.cell(0, 5, f"{company}", 0, 1, 'C')

                    self.set_font('Arial', '', 12)
                    self.cell(0, 5, f"{estate}", 0, 1, 'C')

                    self.set_font('Arial', '', 12)
                    self.cell(0, 5, f"{petak}", 0, 1, 'C')

                    self.set_font('Arial', '', 12)
                    self.cell(0, 5, f"ROTATION {rotasi}", 0, 1, 'C')

                    self.set_font('Arial', '', 12)
                    self.cell(0, 5, f"{start} - {end}", 0, 1, 'C')
                    self.ln(10)
            
                def footer(self):
                    # Add a page number
                    self.set_y(-15)
                    self.set_font('Arial', 'I', 8)
                    self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

                def add_images(self, png_output_report_status, pie_chart, png_output_statistic_date):
                    # Add the images to the PDF
                    self.add_page()

                    # First image: place it on the left side
                    self.image(png_output_report_status, x=5, y=40, w=200)
                    
                    # Second image: place it on the right side
                    self.image(pie_chart, x=20, y=165, w=70)

                    # Fourth image: place it on the right side
                    self.image(png_output_statistic_date, x=100, y=172, w=95)
                    # self.add_page(orientation='L')

                    # self.image(png_output_report_status1, x=5, y=50, w=280)


                def add_table1(sel, dataframe,x ,y, table_title, png_output_report_status1, blok, hektaran):
                    # sel.set_xy(20,220)
                    sel.set_xy(x-30, y+15)
                    sel.set_font('Arial', 'B', 11)
                    sel.cell(0, 8, f'Total Area for Block {blok}', 0, 1, 'L')
                    sel.set_xy(x-30, y+22)
                    sel.cell(0, 8, f'= {hektaran} ha', 0, 1, 'L')
                    
                    sel.set_xy(x, y)
                    sel.set_font('Arial', 'B', 12)
                    sel.cell(0, 10, table_title, 0, 1, 'C')
                    # sel.set_title('Spray Accuration Summary')
                    sel.set_xy(x+25, y + 10)
                    sel.set_font('Arial', 'B', 12)
                    # sel.set_title('Spray Accuration Summary')
                    sel.cell(35, 8, 'Status', 1, 0, 'C')
                    sel.cell(35, 8, 'Number of Trees', 1, 0, 'C')
                    sel.cell(35, 8, 'Percentage', 1, 0, 'C')
                    sel.ln()
                    sel.set_font('Arial', '', 10)
                    # sel.set_xy(20,229)
                    for i in range(len(dataframe)):
                        sel.set_x(x+25)
                        if i == len(dataframe) - 1:  # Check if this is the last row (Total row)
                            sel.set_font('Arial', 'B', 10)  # Set font to bold for the Total row
                        else:
                            sel.set_font('Arial', '', 10)  # Set font to regular for other rosheet_df
                        sel.set_fill_color(200, 220, 255)
                        fill = True
                        sel.cell(35, 6, str(dataframe.iloc[i, 0]), 1, 0, 'C')
                        sel.cell(35, 6, str(dataframe.iloc[i, 1]), 1, 0, 'C')
                        sel.set_fill_color(200, 220, 255)
                        # sel.set_font('Arial', 'B', 10)
                        sel.cell(35, 6, f"{dataframe.iloc[i, 2]}", 1, 0,'C',fill)
                        sel.ln()
                    
                    sel.add_page(orientation='L')
                    sel.image(png_output_report_status1, x=5, y=40, w=280)
                    
                    
                def add_table(sel, dataframe, x, y, table_title):
                    sel.add_page(orientation='P')
                    
                    # sel.set_xy(20,220)
                    sel.set_xy(x+10, y-10)
                    sel.set_font('Arial', 'B', 12)
                    sel.cell(0, 5, table_title, 0, 1, 'C')
                    # sel.set_title('Spray Accuration Summary')
                    sel.set_xy(x + 20, y )
                    sel.set_font('Arial', 'B', 12)
                    # sel.set_title('Spray Accuration Summary')
                    sel.cell(10, 8, 'No', 1, 0, 'C')
                    sel.cell(30, 8, 'X_Pohon', 1, 0, 'C')
                    sel.cell(30, 8, 'Y_Pohon', 1, 0, 'C')
                    sel.cell(30, 8, 'X_Drone', 1, 0, 'C')
                    sel.cell(30, 8, 'Y_Drone', 1, 0, 'C')
                    sel.cell(20, 8, 'Jarak', 1, 0, 'C')
                    sel.cell(25, 8, 'Status', 1, 0, 'C')
                    sel.ln()
                    sel.set_font('Arial', '', 10)
                    # sel.set_xy(20,229)
                    
                    for i in range(len(dataframe)):
                        sel.set_x(x + 20)

                        # Ambil nilai dengan penanganan NaN
                        no = dataframe.iloc[i, 0]
                        x_pohon = dataframe.iloc[i, 1]
                        y_pohon = dataframe.iloc[i, 2]
                        x_drone = dataframe.iloc[i, 3]
                        y_drone = dataframe.iloc[i, 4]
                        jarak = dataframe.iloc[i, 5]
                        status = dataframe.iloc[i, 6]

                        def format_decimal(value, max_decimals=8):
                            if pd.notna(value):
                                return f"{value:.{max_decimals}f}".rstrip('0').rstrip('.')
                            return '-'
                        
                        if pd.isna(jarak):
                            value_str = '-'  # Jika NaN, tampilkan '-'
                        else:
                            value_str = f"{jarak:.2f} cm"  # Jika bukan NaN, tampilkan nilai dengan 'cm'

                        x_drone_str = format_decimal(x_drone)
                        y_drone_str = format_decimal(y_drone)
                        # jarak_str = f"{format_decimal(jarak, 2)} cm"

                        # Tambahkan ke tabel
                        sel.cell(10, 6, str(no), 1, 0, 'C')
                        sel.cell(30, 6, str(x_pohon), 1, 0, 'C')
                        sel.cell(30, 6, str(y_pohon), 1, 0, 'C')
                        sel.cell(30, 6, x_drone_str, 1, 0, 'C')
                        sel.cell(30, 6, y_drone_str, 1, 0, 'C')
                        sel.cell(20, 6, value_str, 1, 0, 'C')

                        # Format warna sel untuk status
                        sel.set_fill_color(200, 220, 255)
                        sel.cell(25, 6, str(status), 1, 0, 'C', fill=True)
                        sel.ln()


            
            self.progressBar.setValue(0)
            dfs = []

            for filename in os.listdir(self.data_pdf):
                if filename.endswith('.csv'):
                    file_path = os.path.join(self.data_pdf, filename)

                    df = pd.read_csv(file_path)
                    dfs.append(df)

            combined_df = pd.concat(dfs, ignore_index=True)
            print(combined_df)
            
            text_20 = str(self.lineEdit_20.text())
            
            
            if 'misi' in combined_df.columns:
                # Filter data yang memiliki misi yang mengandung 'P23A' tanpa memperhatikan huruf besar/kecil
                filtered_df = combined_df[combined_df['misi'].str.contains(text_20, case=False)]
                print("Kolom 'misi' ditemukan, data difilter.")
                # filtered_df.to_csv(os.path.join(self.data_output_pdfmerge,"filter.csv"))
            else:
                # Jika tidak ada kolom 'misi', tidak melakukan filter
                filtered_df = combined_df
                print("Kolom 'misi' tidak ditemukan, data tidak difilter.")
            
            
            filtered_df = filtered_df[['x_penyemprotan', 'y_penyemprotan', 'jarak', 'Status', 'x1', 'y1', 'x2', 'y2', 'misi', 'tanggal']]
            # Menambahkan kolom 'No' menggunakan reset_index
            
            
            status_priority = {"Accurate": 1, 
                                "Not Accurate": 2, 
                            #    "Missed": 3, 
                                "Not Sprayed": 3}
            # Add a priority column to the data
            filtered_df['priority'] = filtered_df['Status'].map(status_priority)
            # Sort data by priority and drop duplicates based on x1 and y1
            filtered_df = filtered_df.sort_values('priority').drop_duplicates(subset=['x1', 'y1'], keep='first')
            # Drop the priority column
            filtered_df = filtered_df.drop(columns=['priority'])
            

            # Save the filtered data to a new file
            # filtered_data.to_csv('C:\\Users\\user\\Downloads\\Q-27\\Q-27\\2\\filtered_data.csv', index=False)

            # print("Data filtering complete. Filtered data saved to 'filtered_data.csv'.")
            
            filtered_df = filtered_df.reset_index(drop=True)
            filtered_df['No'] = filtered_df.index + 1
            
            filtered_df['Status'] = filtered_df['Status'].replace('Missed', 'Not Sprayed')


            # filtered_df.to_csv(os.path.join(self.data_output_pdfmerge,"filter.csv"))
            
            
            statistik_penyemprotan = filtered_df['Status'].value_counts()
            # print(statistik_penyemprotan)
                        
            data_counts = filtered_df['Status'].value_counts()
          
            
            contoh_pie2 = filtered_df[filtered_df['Status'].isin(['Accurate', 'Not Accurate'])]
            data_counts2 = contoh_pie2['Status'].value_counts()
            
            status_color_map = {
            'Accurate': 'limegreen',
            'Not Accurate': 'dodgerblue',
            # 'Missed': 'red',
            'Not Sprayed' : 'brown'
            }

            # Pastikan urutan warna sesuai dengan urutan status di pie chart
            colors = [status_color_map[status] for status in data_counts2.index]
            explode = [0.02] * len (data_counts2)
            plt.pie(data_counts2.values, labels=data_counts2.index, autopct='%0.01f%%', explode=explode, shadow={'ox': -0.01, 'edgecolor': 'none', 'shade': 0.9}, startangle=90, colors= colors)
            # Visualisasi pie chart
            
            png_output_pie_chart = os.path.join(self.data_output_pdfmerge, '2.piechart.png')
            plt.savefig(png_output_pie_chart, format='png', bbox_inches='tight', dpi=200)
            plt.close()
            self.progressBar.setValue(20)
            
            
            df2 = pd.DataFrame()
            df2['Status'] = statistik_penyemprotan.index
            df2['Number of Trees'] = statistik_penyemprotan.values
            
            total_sprayed = df2[df2['Status'].isin(['Accurate', 'Not Accurate'])]['Number of Trees'].sum()
            total_mission = total_sprayed + df2[df2['Status'] == 'Not Sprayed']['Number of Trees'].sum()
            
            
            df2.loc[df2['Status'] == 'Accurate', 'Percentage'] = ((df2.loc[df2['Status'] == 'Accurate', 'Number of Trees'] / total_sprayed * 100).round(2).astype(str) + "%")
            df2.loc[df2['Status'] == 'Not Accurate', 'Percentage'] = ((df2.loc[df2['Status'] == 'Not Accurate', 'Number of Trees'] / total_sprayed * 100).round(2).astype(str) + "%")
            df2.loc[df2['Status'] == 'Not Sprayed', 'Percentage'] = ""
                    
            df2 = pd.concat([df2, pd.DataFrame([
                ['Total Sprayed', total_sprayed, '100 %'],
                ['Total Mission', total_mission, '']
            ], columns=['Status', 'Number of Trees', 'Percentage'])], ignore_index=True)
            
            
            status_order = ["Accurate", "Not Accurate", "Not Sprayed", "Total Sprayed", "Total Mission"]
            df2['Status'] = pd.Categorical(df2['Status'], categories=status_order, ordered=True)
            df2 = df2.sort_values('Status').reset_index(drop=True)
            print(df2)


            self.progressBar.setValue(40)
                
            x_coords = filtered_df['x1']
            y_coords = filtered_df['y1']
            status = filtered_df['Status']
            colors = {'Accurate': 'limegreen', 
                        # 'Missed': 'brown', 
                        'Not Accurate': 'dodgerblue', 
                        'Not Sprayed' : 'brown'}
            color_values = [colors[s] for s in status]

            plt.figure(figsize=(10, 6))
            plt.scatter(x_coords, y_coords, c=color_values, s=3)
            handles = [plt.Line2D([0], [0], marker='o', color='w', markerfacecolor=color, markersize=7) for color in colors.values()]
            labels = list(colors.keys())
            plt.legend(handles, labels)
            # plt.xlabel('Longitude')
            # plt.ylabel('Latitude')
            plt.text(0.465, 1.05, 'Status of Tree', ha='center', va='center', transform=plt.gca().transAxes, fontsize=12, fontweight='bold')
            plt.gca().xaxis.get_major_formatter().set_useOffset(False)
            png_output_report_status = os.path.join(self.data_output_pdfmerge, "1.all_status_report.png")
            plt.savefig(png_output_report_status, format='png', bbox_inches='tight', dpi=200)
            plt.close()


            try:
                filtered_df['tanggal'] = pd.to_datetime(filtered_df['tanggal'], dayfirst=True, errors='coerce')
                if filtered_df['tanggal'].isnull().any():
                    raise ValueError("Beberapa tanggal gagal diproses.")
            except ValueError as ve:
                QMessageBox.warning(None, "ERROR", f"Kesalahan parsing tanggal: {str(ve)}")

            # filtered_df['tanggal'] = pd.to_datetime(filtered_df['tanggal'], format='%d-%B-%Y')

            date_statistic = filtered_df['tanggal'].value_counts().sort_index()
            date_statistic_df = date_statistic.reset_index()
            date_statistic_df.columns = ['tanggal', 'count']
            date_statistic_df = date_statistic_df.sort_values('tanggal')
            
            start = filtered_df['tanggal'].dt.strftime('%d-%B-%Y').min()

            start1 = filtered_df['tanggal'].min()
            start = start1.strftime('%d-%B-%Y')

            end1 = filtered_df['tanggal'].max()
            end = end1.strftime('%d-%B-%Y')
            
            self.progressBar.setValue(60)

            # Plotting the data
            plt.figure(figsize=(12, 6))
            plt.xticks(rotation=30) 
            plt.plot(date_statistic_df['tanggal'], date_statistic_df['count'], color='Blue', linestyle='-', marker='o')
            # Customizing text with bold font and specific size
            plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%d-%B-%Y'))
            plt.xlabel('Date', fontsize=12, fontweight='bold')
            plt.ylabel('Tree', fontsize=12, fontweight='bold')
            plt.title('Tree Statistics by Date', fontsize=16, fontweight='bold')
            plt.xticks(fontsize=14)
            plt.yticks(fontsize=14)
            # Saving the plot
            png_output_statistic_date = os.path.join(self.data_output_pdfmerge, "3.statistic_date.png")
            plt.savefig(png_output_statistic_date, format='png', bbox_inches='tight', dpi=200)
            plt.close()
            
            kml = simplekml.Kml()
            for _, row in filtered_df.iterrows():
                kml.newpoint(
                    name=row['Status'],
                    coords=[(row['x1'], row['y1'])],
                )
            # # Menyimpan file KML
            kml.save(os.path.join(self.data_output_pdfmerge, f"Accuration Report {self.lineEdit_20.text()} {self.comboBox_4.currentText()} Rotation {self.spinBox_1.value()}.kml"))
            
            self.progressBar.setValue(70)

            df3 = pd.DataFrame()

            # Mengatur penomoran otomatis dimulai dari 1 tanpa mengikutkan baris kosong dari filtered_df
            df3['No'] = filtered_df['No']

            # Menambahkan data dari 'filtered_df' ke 'df3'
            df3['X_Pohon'] = filtered_df['x1']
            df3['Y_Pohon'] = filtered_df['y1']
            df3['X_Drone'] = filtered_df['x2']
            df3['Y_Drone'] = filtered_df['y2']
            df3['jarak'] = (filtered_df['jarak'] * 100).round(2)
            df3['status'] = filtered_df['Status']
            
        
            
            x_coords = df3['X_Pohon']
            y_coords = df3['Y_Pohon']
            status = df3['status']
            colors = {'Accurate': 'limegreen', 
                    #   'Missed': 'red', 
                        'Not Accurate': 'dodgerblue', 
                        'Not Sprayed' : 'brown'}
            # color_values = [colors[s] for s in status]

            plt.figure(figsize=(15, 8))
            plt.scatter(x_coords, y_coords, c=color_values, s=3)
            handles = [plt.Line2D([0], [0], marker='o', color='w', markerfacecolor=color, markersize=7) for color in colors.values()]
            for i in range(len(x_coords)):
                plt.text(x_coords[i], y_coords[i], str(i+1), fontsize=3, ha='right')
                
            labels = list(colors.keys())
            plt.legend(handles, labels)
            # plt.xlabel('Longitude')
            # plt.ylabel('Latitude')
            plt.text(0.465, 1.05, 'Status of Tree', ha='center', va='center', transform=plt.gca().transAxes, fontsize=12, fontweight='bold')
            plt.gca().xaxis.get_major_formatter().set_useOffset(False)
            png_output_report_status1 = os.path.join(self.data_output_pdfmerge, "1.all_status_report1.png")
            plt.savefig(png_output_report_status1, format='png', bbox_inches='tight', dpi=200)
            plt.close()
            
            pdf = PDF(format='A4')
            pdf.set_auto_page_break(auto=True, margin=15)
            # Add the image
            pdf.add_images(png_output_report_status, png_output_pie_chart, png_output_statistic_date)
            # pdf.add_images1(png_output_report_status1)
            pdf.add_table1(df2, 50, 230, 'Spray Accuration Summary', png_output_report_status1, blok1, hektaran )
            pdf.add_table(df3, 0, 55, 'Spray Accuration Table')
            pdf_output_all = os.path.join(self.data_output_pdfmerge, f"Accuration Report {self.lineEdit_20.text()} {self.comboBox_4.currentText()} Rotation {self.spinBox_1.value()}.pdf")
            # pdf_output_all = os.path.join(self.data_output_pdfmerge,"4_first_page_report_acuracy.pdf")
            pdf.output(pdf_output_all)
            self.progressBar.setValue(100)
            self.progressBar.setFormat("Completed")


                    # menghapus file yang tidak diperlukan
            for file_path in [png_output_report_status1, png_output_report_status, png_output_statistic_date, png_output_pie_chart]:
                if os.path.exists(file_path):
                        os.remove(file_path)
                            

        except Exception as e:
            QMessageBox.critical(None, "Error", f"Terjadi kesalahan yang tidak diketahui: {str(e)}")

        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())