import xlwings as xw
book = xw.Book(r"C:/ADMAC-Parts/FilesM/Output/G001/T1.xlsm")
sht = xw.Book().sheets[0]
sht.range('A1').value = 1