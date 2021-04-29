import xlwings as xw
wb = xw.Book(r"C:/ADMAC-Parts/FilesM/Output/G001/T1.xlsm")
sheet = wb.sheets['Gcode']
sheet.range('A5').value = 'From Script'

f = open("C:/ADMAC-Parts/FilesM/Output/G001/G001.MIN", "r")
#print(f.read())
coding = f.read()
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

wb.save(r"C:/ADMAC-Parts/FilesM/Output/G001/T1.xlsm")