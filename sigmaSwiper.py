#!/usr/bin/env python3
import sys
import os
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, qApp
import pandas as pd
import datetime
import openpyxl
from sigmaSwiperGui import Ui_sigmaSwiper
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import configparser
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import(
        FigureCanvasQTAgg as FigureCanvas,
        )
import matplotlib.dates as mdates

class SigmaSwiperProgram(QtWidgets.QMainWindow,Ui_sigmaSwiper):
    
    guest_list = {"ID":[],
                "NAME":[]}
    has_guest = False
    has_graph = False
    data = {"TIME":[],
            "ID":[],
            "NAME":[]}
    graph_x = []
    graph_y = []
    today=datetime.datetime.now().strftime("%m-%d-%y")
    count = 0
    settings = {}
    settings_file=".settings.ini"
    config = configparser.ConfigParser()
    fig = plt.Figure(figsize=(6,4))
    def __init__(self):
        Ui_sigmaSwiper.__init__(self)
        QtWidgets.QMainWindow.__init__(self)
        self.setupUi(self)
        self.submit_button.clicked.connect(self.read_ID)
        self.action_export_list.triggered.connect(self.export_data)
        self.action_load_guest_list.triggered.connect(self.input_guest_list)
        self.action_exit.triggered.connect(qApp.quit)
        self.config.read(self.settings_file)
        self.settings = self.config['settings']
        self.guest_list_check_label.setText("Load a guest list")
        self.id_input.returnPressed.connect(self.read_ID)
        self.graph_layout = QtWidgets.QVBoxLayout()
        self.graph_widget.setLayout(self.graph_layout)
    def input_guest_list(self):
        fname = QFileDialog.getOpenFileName(None, 'Open Guestlist' , os.path.expanduser('~')+"/Desktop/", "Excel Files (*.xlsx)")
        if fname[0] == '': 
            pass
        else:
            excel_file = pd.ExcelFile(fname[0])
            df = excel_file.parse("Sheet1")
            self.guest_list["ID"] = df["ID"].tolist()   
            self.guest_list["NAME"] = df["NAME"].tolist()
            self.has_guest = True
            self.guest_list_check_label.setStyleSheet('color: white')
            self.guest_list_check_label.setText("Guest List Loaded")
    def read_ID(self):
        inp = self.id_input.text()
        if len(inp) == 6:
            if self.has_guest:
                if int(inp) not in self.guest_list["ID"]:
                    self.guest_list_check_label.setText("Not On List")
                    self.guest_list_check_label.setStyleSheet('color: red')
                    self.id_input.clear()
                else:
                    self.count +=1
                    self.graph_y.append(self.count)
                    self.data["ID"].append(inp)
                    self.data["TIME"].append(datetime.datetime.now().strftime("%H:%M"))
                    self.graph_x.append(datetime.datetime.now().strftime("%H:%M:%S"))
                    name =self.guest_list["NAME"][self.guest_list["ID"].index(int(inp))]
                    self.data["NAME"].append(name)
                    self.guest_list_check_label.setText("On List")
                    self.guest_list_check_label.setStyleSheet('color: green')
                    self.lcdNumber.display(self.count)
                    self.list_preview.addItem(name+" - "+inp)
                    self.plot_data()
            else:
                self.guest_list_check_label.setText("No Guest List Loaded")
                self.guest_list_check_label.setStyleSheet('color: yellow')
                self.id_input.clear()
            
        elif len(inp) == 13:
            inp = inp[4:10]
            if self.has_guest:
                if int(inp) not in self.guest_list["ID"]:
                    self.guest_list_check_label.setText("Not On List")
                    self.guest_list_check_label.setStyleSheet('color: red')
                    self.id_input.clear()
                else:
                    self.count +=1
                    self.graph_y.append(self.count)
                    self.data["ID"].append(inp)
                    self.data["TIME"].append(datetime.datetime.now().strftime("%H:%M"))
                    self.graph_x.append(datetime.datetime.now().strftime("%H:%M:%S"))
                    name =self.guest_list["NAME"][self.guest_list["ID"].index(int(inp))]
                    self.data["NAME"].append(name)
                    self.guest_list_check_label.setText("On List")
                    self.guest_list_check_label.setStyleSheet('color: green')
                    self.lcdNumber.display(self.count)
                    self.list_preview.addItem(name+" - "+inp)
                    self.plot_data()
            else:
                self.guest_list_check_label.setText("No Guest List Loaded")
                self.guest_list_check_label.setStyleSheet('color: yellow')
                self.id_input.clear()
        else:
            self.guest_list_check_label.setText("Invalid Input")
            self.guest_list_check_label.setStyleSheet('color: yellow')
        self.id_input.clear()
    
    def email_list(self,file_path):
        if self.settings["send_email"] == "yes":
            for x in self.settings["to_email"].strip().split(','):
                try:
                    toaddr = x.strip()
                    fromaddr = self.settings["from_email"]
                     
                    msg = MIMEMultipart()
                      
                    msg['From'] = fromaddr
                    msg['To'] = toaddr
                    msg['Subject'] = "Party List for"+self.today
                       
                    body = "Hello,\n Attached is the attendance sheet for our social event on "+self.today+". If any additional information is needed, please contact <insert responsible person here>"
                        
                    msg.attach(MIMEText(body, 'plain'))
                    filename = file_path.split("/")[-1]
                    attachment = open(file_path, "rb")
                          
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload((attachment).read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
                           
                    msg.attach(part)
                            
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(fromaddr, self.settings["email_password"])
                    text = msg.as_string()
                    server.sendmail(fromaddr, toaddr, text)
                    server.quit()
                except:
                    print("email error")
                    pass
        else:
            pass
    
    def export_data(self):
        export = pd.DataFrame(self.data)        
        export.index += 1
        fname = QFileDialog.getSaveFileName(None, 'Save Guest Log' , os.path.expanduser('~')+"/Desktop/"+self.today+"-"+self.settings["default_filename"]+".xlsx","Excel Files (*.xlsx)" )
        if fname[0] == '':
            pass
        else:
            export.to_excel(fname[0])
        self.email_list(fname[0])
     
    def plot_data(self):
        if not self.has_graph:
            new_x = [mdates.datestr2num(x) for x in self.graph_x]
            plt.xlabel('Time')
            plt.ylabel('Count')
            ax = self.fig.add_subplot(111)
            ax.plot(new_x, self.graph_y, 'r-')
            #ax.get_xaxis().set_ticks([])
            ax.get_yaxis().set_ticks([])
            self.canvas = FigureCanvas(self.fig)
            self.graph_layout.addWidget(self.canvas)
            self.canvas.draw()
            self.has_graph = True
        else:
            self.graph_layout.removeWidget(self.canvas)
            self.canvas.close()
            self.fig.clf()
            plt.xlabel('Time')
            plt.ylabel('Count')
            ax = self.fig.add_subplot(111)
            new_x = [mdates.datestr2num(x) for x in self.graph_x]
            ax.plot(new_x, self.graph_y, 'r-')
            ax.get_xaxis().set_ticks([])
            #ax.get_yaxis().set_ticks([])
            self.canvas = FigureCanvas(self.fig)
            self.graph_layout.addWidget(self.canvas)
            self.canvas.draw()

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window=SigmaSwiperProgram()
    window.show()
    sys.exit(app.exec_())
