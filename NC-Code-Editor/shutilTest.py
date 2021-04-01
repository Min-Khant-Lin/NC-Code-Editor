import os
import shutil

path = 'C:/ADMAC-Parts/FilesM/Output'

print("Before copying file:")
print(os.listdir(path))

source = "C:/ADMAC-Parts/FilesM/Output/NCコード-V13.xlsm"

perm = os.stat(source).st_mode
print("File Permission mode:", perm, "/n")

fileName = "G001"

destination = "C:/ADMAC-Parts/FilesM/Output/" + fileName + "/" + fileName + ".xlsm"

# Copy from source to destination
dest = shutil.copy(source, destination)

print(dest)