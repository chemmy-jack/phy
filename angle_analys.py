# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

# -*- coding: utf-8 -*-
"""
Created on Sat Aug 10 23:21:55 2019

@author: User
"""

import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt


file_name = "0831_4_1"

file = xw.Book("/Volumes/jack32/physic/motion_analysis/"+file_name+"/"+file_name+".xlsx")


# sht = file.sheets[0]
# 引用單元格
# rng = sht.range('a1')
# rng = sht['a1']
# rng = sht[0,0] 第一行的第一列即a1,相當於pandas的切片
# 引用區域
# rng = sht.range('a1:a5')
# rng = sht['a1:a5']
# rng = sht[:5,0]

nrows = file.sheets[2]["a1"].expand("table").rows.count - 2

s = 2180/800 #to sincronize x of top(x:1280,y:800) and side(x:800,y:600)


data = file.sheets["data"]
side_wb = np.array(data[f""+chr(96+1)+"3:"+chr(96+2)+str(nrows+2)].value)  # wingbase
side_wt = np.array(data[f""+chr(96+3)+"3:"+chr(96+4)+str(nrows+2)].value)  # wingtip
side_te = np.array(data[f""+chr(96+5)+"3:"+chr(96+6)+str(nrows+2 )].value)  # hindwing trailing edge
side_tail = np.array(data[f""+chr(96+7)+"3:"+chr(96+8)+str(nrows+2)].value)

top_wb = np.array(data[f""+chr(96+9)+"3:"+chr(96+10)+str(nrows+2)].value)*s  # wingbase
top_wt = np.array(data[f""+chr(96+11)+"3:"+chr(96+12)+str(nrows+2)].value)*s  # wingtip
top_te = np.array(data[f""+chr(96+13)+"3:"+chr(96+14)+str(nrows+2)].value)*s  # hindwing trailing edge
top_tail = np.array(data[f""+chr(96+15)+"3:"+chr(96+16)+str(nrows+2)].value)*s
print("finnish colecting data and multiply top with 2180/800")
#multiply top with 2180/800




head = []
tail= []
wb = []
wt = []
te = []

# offset
side_offset = [side_wb[0][0],side_wb[0][1]]
top_offset = [top_wb[0][0],top_wb[0][1]]


for i in range(nrows):
    for j in range(2):
       # side_head[i][j] -= side_offset[j]
        side_wt[i][j] -= side_offset[j]
        side_wb[i][j] -= side_offset[j]
        side_te[i][j] -= side_offset[j]
        side_tail[i][j] -= side_offset[j]

       # top_head[i][j] -= top_offset[j]
        top_wt[i][j] -= top_offset[j]
        top_wb[i][j] -= top_offset[j]
        top_te[i][j] -= top_offset[j]      
        top_tail[i][j] -= top_offset[j]


    #head.append( [-(side_head[i][0] + top_head[i][0]) / 2, -side_head[i][1], -top_head[i][1]])
    wb.append( [-(side_wb[i][0] + top_wb[i][0]) / 2, -side_wb[i][1], -top_wb[i][1]])
    wt.append( [-(side_wt[i][0] + top_wt[i][0]) / 2, -side_wt[i][1], -top_wt[i][1]])
    te.append( [-(side_te[i][0] + top_te[i][0]) / 2, -side_te[i][1], -top_te[i][1]])
    tail.append([-(side_tail[i][0] + top_tail[i][0]) / 2, -side_tail[i][1], -top_tail[i][1]])
print("1/2")
#head = np.array(head)
wb = np.array(wb)
wt = np.array(wt)
te = np.array(te)
tail = np.array(tail)


body_vector = wb - tail
le_vector = wt - wb
hind_te_vector = te - wb

sweeping_angle = []
pitching_angle = []
flapping_angle = []


for i in range(nrows):

    wingplane_normal_vector = np.cross(le_vector[i], hind_te_vector[i])
    sw_base_vector = np.cross(wingplane_normal_vector, body_vector[i])
    len_sw_base = np.sqrt(np.dot(sw_base_vector, sw_base_vector))
    len_le = np.sqrt(np.dot(le_vector[i], le_vector[i]))
    sw_temp = np.arccos(np.dot(sw_base_vector, le_vector[i]) / (len_sw_base * len_le)) * 180 / np.pi
    sweeping_angle.append(sw_temp)

    pitching_angle.append(np.arctan(body_vector[i][1] / body_vector[i][0]) * 180 / np.pi)

    body_right_temp_x = body_vector[i][2] * np.sqrt(body_vector[i][0] ** 2 + body_vector[i][2] ** 2)
    body_right_temp_z = -body_vector[i][0] * np.sqrt(body_vector[i][0] ** 2 + body_vector[i][2] ** 2)
    # body_right_vector = np.array([body_right_temp_x,wb[i][1],body_right_temp_z])
    body_right_vector = np.array([body_right_temp_x, 0, body_right_temp_z])

    len_body_right = np.sqrt(np.dot(body_right_vector, body_right_vector))

    flap_temp = np.dot(sw_base_vector, body_right_vector) / (len_sw_base * len_body_right)
    #flap_temp = np.dot(sw_base_vector, [0, 0, -1]) / (len_sw_base)

    if sw_base_vector[1] > 0:
        flap_temp = np.arccos(flap_temp) * 180 / np.pi
    else:
        flap_temp = -np.arccos(flap_temp) * 180 / np.pi

    flapping_angle.append(-flap_temp)

print("2/2")

#create angle sheet
try :
    print(file.sheets["angle"])
except :
    file.sheets.add(name="angle",after=-1)
    print(file.sheets["angle"])
angle = file.sheets["angle"]
#create number
for i in range(1,nrows+1):
    angle["a"+str(i+1)].value = str(i)
print("side bar number create finnish")
#flapping_angle 
angle["b"+str(1)].value = "flapping angle"
for i in range(1,nrows+1):
    angle["b"+str(i+1)].value = str(flapping_angle[i-1])
print("flapping angle sheet finnished")
# pitching_angle 
angle["c"+str(1)].value = "pitching angle"
for i in range(1,nrows+1):
    angle["c"+str(i+1)].value = str(pitching_angle[i-1])
print("pitching angle sheet finnished")
# sweeping_angle 
angle["d"+str(1)].value = "sweeping angle"
for i in range(1,nrows+1):
    angle["d"+str(i+1)].value = str(sweeping_angle[i-1])
print("sweeping angle sheet finnished")



temp = np.arange(0, nrows)
plt.plot(temp, pitching_angle,label = "pitching")
plt.plot(temp, flapping_angle,label = "flapping")
plt.plot(temp,sweeping_angle,label = "sweeping")
plt.legend()
plt.show()
