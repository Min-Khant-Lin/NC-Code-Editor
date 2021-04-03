import openpyxl

xfile = openpyxl.load_workbook(filename = "C:/ADMAC-Parts/FilesM/Output/G001/T1.xlsm", read_only=False, keep_vba=True)

xsheet = xfile["Gcode"]
f = open("C:/ADMAC-Parts/FilesM/Output/G001/G001.MIN", "r")
#print(f.read())
coding = f.read()
oneLine = ""

i = 2
for x in coding:
    if x == "\n":
        xsheet['A' + str(i)] = oneLine
        oneLine = ""
        i += 1
    else:
        oneLine+=x
    
i = 0
xfile.save("C:/ADMAC-Parts/FilesM/Output/G001/T1.xlsm")
#commentbyTDH
