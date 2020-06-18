import tkinter
import tkinter.filedialog

files = [('All Files', '*'), ('Excel 2010', '*.xlsx'), ('Comma Split Values', '*.csv')]


def getOpenDir():
    root = tkinter.Tk()
    root.withdraw()
    file = tkinter.filedialog.askopenfilename(title="Open Excel Spreadsheet", filetypes=files, defaultextension=files)
    root.destroy()
    return file


def getSaveDir():
    root = tkinter.Tk()
    root.withdraw()
    file = tkinter.filedialog.asksaveasfilename(title="Save Excel Spreadsheet", filetypes=files, defaultextension=files)
    root.destroy()
    return file


if __name__ == '__main__':
    print(getOpenDir())
