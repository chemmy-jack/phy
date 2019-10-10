import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt


file_name = "test"

file = xw.Book("/Volumes/jack32/motion_analysis/"+file_name+"/"+file_name+".xlsx")
print(file)
print(file.sheets[0]["a1"].value)
n = 10

file.sheets[0][f"a1:"+chr(96+n)+str(n)].value = "1234"


for i in range(1,n+1) :
    print(i)
    file.sheets[0][f""+chr(96+i)+str(i)+":"+chr(96+i)+str(i)+""].value = str(i)



