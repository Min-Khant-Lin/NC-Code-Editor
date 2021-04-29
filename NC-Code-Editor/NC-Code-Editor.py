import tkinter as tk
from tkinter import messagebox

def createExcel():
    pName = entry1.get()
    print(pName)
    try:
        with open("C:/ADMAC-Parts/FilesM/Output/" + pName + "/" + pName + ".MIN") as f:
            contents = f.read()
            print(contents)
    except (OSError,FileNotFoundError):
        messagebox.showerror("エラー", pName + "が見つかりませんでした。")

window = tk.Tk()

w = 200 #tkの長さ
h = 150 #tkの高さ
ws = window.winfo_screenwidth() #widht of screen
hs = window.winfo_screenheight()

window.title("M.K.L")
window.geometry("300x150")

frame1 = tk.Frame()

label1 = tk.Label(master = frame1, text = "プログラム名を記入してください。")
label1.pack()

entry1 = tk.Entry(frame1, width = 20)
entry1.pack(padx =5, pady =5)

eButton = tk.Button(frame1, text = "エクセル作成", command = createExcel)
eButton.pack()

frame1.pack()
window.mainloop()