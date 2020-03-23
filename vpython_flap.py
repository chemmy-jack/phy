import xlwings as xw
import vpython as vp
import numpy as np


file_name = "0902_15_5"

fil = xw.Book("/Volumes/jack32/physic/motion_analysis/"+file_name+"/"+file_name+"(unit_cm).xlsx")
nrows = fil.sheets[2]["a1"].expand("table").rows.count - 2

data = fil.sheets["data"]
side_wb = np.array(data[f""+chr(96+1)+"3:"+chr(96+2)+str(nrows+2)].value)  # wingbase
side_wt = np.array(data[f""+chr(96+3)+"3:"+chr(96+4)+str(nrows+2)].value)  # wingtip
side_te = np.array(data[f""+chr(96+5)+"3:"+chr(96+6)+str(nrows+2 )].value)  # hindwing trailing edge
side_tail = np.array(data[f""+chr(96+7)+"3:"+chr(96+8)+str(nrows+2)].value)

top_wb = np.array(data[f""+chr(96+9)+"3:"+chr(96+10)+str(nrows+2)].value)  # wingbase
top_wt = np.array(data[f""+chr(96+11)+"3:"+chr(96+12)+str(nrows+2)].value)  # wingtip
top_te = np.array(data[f""+chr(96+13)+"3:"+chr(96+14)+str(nrows+2)].value)  # hindwing trailing edge
top_tail = np.array(data[f""+chr(96+15)+"3:"+chr(96+16)+str(nrows+2)].value)

head = []
tail= []
wb = []
wt = []
te = []

for i in range(nrows):
    wb.append( [-(side_wb[i][0] + top_wb[i][0]) / 2, -side_wb[i][1], -top_wb[i][1]])
    wt.append( [-(side_wt[i][0] + top_wt[i][0]) / 2, -side_wt[i][1], -top_wt[i][1]])
    te.append( [-(side_te[i][0] + top_te[i][0]) / 2, -side_te[i][1], -top_te[i][1]])
    tail.append([-(side_tail[i][0] + top_tail[i][0]) / 2, -side_tail[i][1], -top_tail[i][1]])
wb = np.array(wb)
wt = np.array(wt)
te = np.array(te)
tail = np.array(tail)
wtr = []
ter = []


