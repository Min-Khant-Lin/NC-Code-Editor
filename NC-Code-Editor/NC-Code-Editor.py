#2021/07/12
import tkinter as tk
from tkinter import messagebox
import os 
import shutil
import xlwings as xw

def createExcel():
    pName = entry1.get()
    try:
        with open("C:/ADMAC-Parts/FilesM/Output/" + pName + "/" + pName + ".MIN") as f:
            contents = f.read()
            #print(contents)
            copyExcel(pName)
            copyCode(contents, pName)
    except (OSError,FileNotFoundError):
        messagebox.showerror("エラー", pName + "が見つかりませんでした。")

def copyExcel(pName):
    path = 'C:/ADMAC-Parts/FilesM/Output'
    source = os.getcwd() + "/Data/NCコード-V13.xlsm"     
    
    perm = os.stat(source).st_mode

    destination = "C:/ADMAC-Parts/FilesM/Output/" + pName + "/" + pName + ".xlsm"
    dest = shutil.copy(source, destination)

def copyCode(contents, pName):
    wb = xw.Book(r"C:/ADMAC-Parts/FilesM/Output/" + pName + "/" + pName + ".xlsm")
    sheet = wb.sheets['Gcode']
    coding = contents
    oneLine = ""

    i = 2
    for x in coding:
        if x == "\n":
            sheet.range('A' + str(i)).value = oneLine
            oneLine = ""
            i += 1
        else:
            oneLine+=x
            if oneLine == "RTS":
                sheet.range('A' + str(i)).value = oneLine
        
    i = 0

    wb.save(r"C:/ADMAC-Parts/FilesM/Output/" + pName + "/" + pName + ".xlsm")

window = tk.Tk()

w = 200
h = 150
ws = window.winfo_screenwidth()
hs = window.winfo_screenheight()
x = (ws / 2) - (w / 2)
y = (hs / 2) - (h / 2)

window.title("M.K.L")
window.geometry("%dx%d+%d+%d" % (w, h, x, y))

frame1 = tk.Frame()

label1 = tk.Label(master = frame1, text = "プログラム名を記入してください。")
label1.pack()

entry1 = tk.Entry(frame1, width = 20)
entry1.pack(padx =5, pady =5)

eButton = tk.Button(frame1, text = "エクセル作成", command = createExcel)
eButton.pack()

frame1.pack()
window.mainloop()