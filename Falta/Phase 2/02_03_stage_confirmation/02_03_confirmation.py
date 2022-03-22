# -*- coding: utf-8 -*-
"""
Created on Mon Nov 22 10:36:27 2021

@author: deep.g
"""

from tkinter import Tk
from tkinter import Label
from tkinter import Entry
from tkinter import Button
from tkinter import Frame
from tkinter import END
from os.path import join
from time import sleep
from tkinter.messagebox import showerror

from datetime import datetime
from win32com.client import GetObject
from win32com.client import CDispatch
from subprocess import Popen
from subprocess import call
from ast import literal_eval

from cryptography.fernet import Fernet

with open(b"bin/v.key", "rb") as keygen:
    key = keygen.readline()
fernet = Fernet(key)
# print(key)
with open(b"bin/v.lnc", "rb") as file:
    encdata = file.readline()
# print(encdata)
decoded_data = fernet.decrypt(encdata).decode()
validation_date = datetime.strptime(decoded_data, "%Y-%m-%d")
# print(validation_date)


monitoring_time_delay = 3


class SapLogin():
    """
    DESCRIPTION
    SAP automation module for python
    ================================
    Saplogin is a module which will help to automate the SAP related manual workload.


    USE:
        1. create a object of SapLogin class with correct argument-
            $sapObject = SapLogin(sap_id, sap_pw, sap_connection)
        2. After creating the object, user need to login.-
            $sapObject.login()
        3. After login we can set the flowchart of multiple activities-
            for example we want to confirm '05 Stage':
            $sapObject.open_zscan() ------> open the zscan window,
            $sapObject.stage_confirmation_05(module_id="50984324") ------>confirmation in 05 stage window.
            $sapObject.back(1) -------> back to it's previous window
    """

    def __init__(self, sap_id, sap_pw, sap_connection):

        self.sap_id = sap_id
        self.sap_pw = sap_pw
        self.sap_connection = sap_connection
        self._sapPath_ = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        self._timeDelay_ = 1

    def login(self):
        """With this function we can login in SAP"""
        Popen(self._sapPath_)
        sleep(5)
        SapGuiAuto = GetObject("SAPGUI")

        if not type(SapGuiAuto) == CDispatch:
            return
        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == CDispatch:
            self.SapGuiAuto = None
            return

        connection = application.OpenConnection(self.sap_connection, True)

        if not type(connection) == CDispatch:
            self.SapGuiAuto = None
            self.application = None
            return

        self.session = connection.Children(0)

        if not type(self.session) == CDispatch:
            self.SapGuiAuto = None
            self.application = None
            #self.session = None
            return
        self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.sap_id
        self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.sap_pw
        self.session.findById("wnd[0]").sendVKey(0)

    def open_zscan(self):
        """using this function we can open zscan window"""
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "zscan"
        self.session.findById("wnd[0]").sendVKey(0)
        # self.session.findById("wnd[0]").sendVKey(8)

    def stage_confirmation_05(self, module_id):
        """After opening the zscan window """
        self.open_zscan()
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById(
            "wnd[0]/usr/subSUB2:ZBARCODE_SCANNER_COPY:0902/txtP_SN").text = module_id
        self.session.findById(
            "wnd[0]/usr/subSUB2:ZBARCODE_SCANNER_COPY:0902/txtP_SN").caretPosition = 6
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]").sendVKey(11)
        self.session.findById("wnd[0]").sendVKey(0)
        self.back(2)

    def stage_confirmation_10(self, module_id):
        self.open_zscan()
        call("cmd /c stage_10_selection.vbs")
        self.session.findById(
            "wnd[0]/usr/subSUB3:ZBARCODE_SCANNER_COPY:0903/txtP_SN1").text = module_id
        self.session.findById(
            "wnd[0]/usr/subSUB3:ZBARCODE_SCANNER_COPY:0903/txtP_SN1").caretPosition = 6
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]").sendVKey(11)
        self.back(2)

    def stage_confirmation_20(self, module_id):
        self.open_zscan()
        call("cmd /c stage_20_selection.vbs")
        self.session.findById(
            "wnd[0]/usr/subSUB5:ZBARCODE_SCANNER_COPY:0904/txtP_SN2").text = module_id
        self.session.findById(
            "wnd[0]/usr/subSUB5:ZBARCODE_SCANNER_COPY:0904/txtP_SN2").caretPosition = 6
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]").sendVKey(11)
        self.back(2)

    def stage_confirmation_30(self, module_id):
        self.open_zscan()
        call("cmd /c stage_30_selection.vbs")
        self.session.findById(
            "wnd[0]/usr/subSUB6:ZBARCODE_SCANNER_COPY:0906/txtP_SN4").text = module_id
        self.session.findById(
            "wnd[0]/usr/subSUB6:ZBARCODE_SCANNER_COPY:0906/txtP_SN4").caretPosition = 6
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]").sendVKey(11)
        self.back(2)

    def back(self, number_of_time=1):
        for _ in range(0, number_of_time):
            self.session.findById("wnd[0]").sendVKey(3)


def read_config(config_path):
    with open(config_path, "r") as file:
        data = file.read()
    return literal_eval(data)


config_path = "config.txt"
config_data = read_config(config_path)
sap_id = config_data["sap_id"]
sap_pw = config_data["sap_pw"]
data_saving_path = config_data["data_file_path"]
sap_connection = config_data["sap_connection"]
all_framing_line = config_data["all_framing_line"]
_current_framing_line = config_data["current_framing_line"]
timedelay = config_data["timedelay"]
sapSession = SapLogin(sap_id, sap_pw, sap_connection)
sapSession.login()


def submit_confirmation(event=None):
    present_time = datetime.now()
    print(validation_date)
    if validation_date.date() >= present_time.date():
        module_id = module_input_entry.get()

        timestamp = present_time.strftime("%y/%m/%d %H:%M:%S")
        sapSession.stage_confirmation_05(module_id)
        sleep(timedelay)
        sapSession.stage_confirmation_10(module_id)
        sleep(timedelay)
        sapSession.stage_confirmation_20(module_id)
        sleep(timedelay)
        sapSession.stage_confirmation_30(module_id)

        data_log = "{},{},{}".format(
            module_id, _current_framing_line, timestamp)
        data_for_label = "Module Id-{} Line-{} Timestamp-{}".format(module_id,
                                                                    _current_framing_line,
                                                                    timestamp)

        data_saving_file = join(data_saving_path,
                                _current_framing_line,
                                present_time.strftime("%y%m%d.txt"))
        with open(data_saving_file, "a", newline="") as file:
            file.write(data_log+"\n")

        print(data_log)
        datalog_label = Label(frame_datashow, text=data_for_label)
        datalog_label.grid(row=0, column=0)

        module_input_entry.delete(0, END)

    else:
        showerror(
            "Error", "Error Code 101, Validation Related issue. Please contact with Deep Gupta(Phone No: +91-9153436296).!!!!")


root = Tk()
icon_img = r"img/logo.ico"
root.title("20-30-Stage Confirmation")
root.iconbitmap(icon_img)

frame_input = Frame(root)
frame_input.grid(row=0, column=0)

frame_datashow = Frame(root)
frame_datashow.grid(row=1, column=0)

module_input_label = Label(
    frame_input, text="Scan Module ID:", borderwidth=2, relief="groove")
module_input_label.grid(row=0, column=0)


module_input_entry = Entry(frame_input)
module_input_entry.grid(row=0, column=1)

module_submit_btn = Button(frame_input, text="Submit",
                           command=submit_confirmation)
module_submit_btn.grid(row=0, column=2)

module_input_entry.bind('<Return>', submit_confirmation)

root.mainloop()
