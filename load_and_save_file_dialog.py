import tkinter
import tkinter.filedialog

files = [('Excel 2010', '*.xlsx')]


def getOpenDir():
    tkinter.Tk().withdraw()
    file = tkinter.filedialog.askopenfilename(title="Open Excel Spreadsheet", filetypes=files, defaultextension=files)
    return file


def getSaveDir():
    tkinter.Tk().withdraw()
    file = tkinter.filedialog.asksaveasfilename(title="Save Excel Spreadsheet", filetypes=files, defaultextension=files)
    return file



