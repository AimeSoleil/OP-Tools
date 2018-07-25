# -*- coding:utf-8 -*-

from tkinter import *
from tkinter import filedialog
from csv_daterangesplit import *
import os
import platform
import datetime

class App:
    def __init__(self):
        self.root = Tk()
        self.root.title("Ring&Marina")
        center_window(self.root, 580, 512)
        self.root.maxsize(580, 512)
        self.root.minsize(580, 512)

        # input button
        self.button_input = Button(self.root, text="Select Input", width=10, height=2, command=self.openInputFile)
        self.button_input.grid(sticky=W, row=0)
        self.entry_input = Entry(self.root, width=50)
        self.entry_input.grid(sticky=W, row=0, column=1, columnspan=2)

        # output button
        self.button_output = Button(self.root, text="Select Output", width=10, height=2, command=self.selectOutputDir)
        self.button_output.grid(sticky=W, row=1)
        self.entry_output = Entry(self.root, width=50)
        self.entry_output.grid(sticky=W, row=1, column=1, columnspan=2)

        # process/help/about button
        self.button_process = Button(self.root, text="Process", width=10, height=2, command=self.process)
        self.button_process.grid(sticky=W, row=2)
        self.button_help = Button(self.root, text="Help", width=10, height=2, command=self.help)
        self.button_help.grid(sticky=E, row=2, column=1)
        self.button_about = Button(self.root, text="About", width=10, height=2, command=self.about)
        self.button_about.grid(sticky=E, row=2, column=2)

        # log text
        self.text_log = Text(self.root)
        self.text_log.grid(row=3, columnspan=3)

        # open window
        mainloop()

    def openInputFile(self):
        filePath = filedialog.askopenfilename(title='打开文件', filetypes=[('xlsx', '*.xlsx')])
        self.entry_input.delete(0, END)
        self.entry_input.insert(END, filePath)
        self.text_log.insert(END, "{0}Input: {1}{2}".format(getCurrentTime(), filePath, lineBreak()))

    def selectOutputDir(self):
        dirPath = filedialog.askdirectory(title="选择输出目录")
        self.entry_output.delete(0, END)
        self.entry_output.insert(END, dirPath)
        self.text_log.insert(END, "{0}Output: {1}{2}".format(getCurrentTime(), dirPath, lineBreak()))

    def process(self):
        self.text_log.insert(END, "{0}processing...{1}".format(getCurrentTime(), lineBreak()))
        inputPath = self.entry_input.get()
        outputPath = self.entry_output.get()

        if inputPath.strip(' ') is '':
            self.text_log.insert(END, "{0}Warning! Input is null{1}".format(getCurrentTime(), lineBreak()))
            return

        if outputPath.strip(' ') is '':
            self.text_log.insert(END, "{0}Warning! Output is null{1}".format(getCurrentTime(), lineBreak()))
            return

        # try:
        #     output_xlsx = process(csv_from_excel(input=inputPath), temp_output=os.path.join(outputPath, "output.csv"))
        #     self.text_log.insert(END, "{0}Completed! See report {1}{2}".format(getCurrentTime(), output_xlsx, lineBreak()))
        # except Exception as e:
        #     self.text_log.insert(END, "{0}{1}: 请确保Excel内的数据格式合法!!!".format(getCurrentTime(), e))
        output_xlsx = process(csv_from_excel(input=inputPath), temp_output=os.path.join(outputPath, "output.csv"))
        self.text_log.insert(END, "{0}Completed! See report {1}{2}".format(getCurrentTime(), output_xlsx, lineBreak()))

    def help(self):
        help_log = '''
            ============================HELP===========================
            1. Click "Select Input" button and choose your excel file.
            2. Click "Select Output" button and choose where your want
               to save your output excel file.
            3. click Process and wait for report.
            ===========================================================
        '''
        self.text_log.insert(END, help_log)

    def about(self):
        about_log = '''
            ============================ABOUT==========================
            Author: Julian
            Intention: This APP is special for Ring & Marina.
                       Best wishes to you.
                                            2017年08月27日20:59:21
            ===========================================================
        '''
        self.text_log.insert(END, about_log)


def getCurrentTime():
    now = datetime.datetime.now()
    nowstr = "[{}]:".format(now.strftime("%Y-%m-%d %H:%M:%S"))
    return nowstr


def lineBreak():
    line_break_dic = {"Windows": "\r\n", "Linux": "\n", "Darwin": "\n"}
    return line_break_dic[platform.system()]


def get_screen_size(window):
    return window.winfo_screenwidth(), window.winfo_screenheight()


def get_window_size(window):
    return window.winfo_reqwidth(), window.winfo_reqheight()


def center_window(root, width, height):
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width)/2, (screenheight - height)/2)
    root.geometry(size)


if __name__ == '__main__':
    app = App()
