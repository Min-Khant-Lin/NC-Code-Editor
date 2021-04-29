import tkinter as tk

pName = ""

def createExcel():
    pName = entry1.get()
    #print(pName)
    try:
        with open("C:/ADMAC-Parts/FilesM/Output/" + pName + "/" + pName + ".MIN") as f:
            contents = f.read()
            print(contents)
    except (OSError,FileNotFoundError):
        print("Could not find the file:" + pName)

window = tk.Tk()
window.title("編集")
window.geometry("200x150")

frame1 = tk.Frame()

label1 = tk.Label(master = frame1, text="プログラム名を記入してください。")
label1.pack()

entry1 = tk.Entry(frame1, width=20)
entry1.insert(0,"")
entry1.pack(padx=5, pady=5)

eButton = tk.Button(frame1, text = "エクセル", command = createExcel)
eButton.pack()

frame1.pack()

window.mainloop()