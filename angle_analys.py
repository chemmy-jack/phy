# -*- coding: utf-8 -*-
"""
Created on Sat Aug 10 23:21:55 2019

@author: User
"""

import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt
from numpy import linalg as LA

file_name = "0322 3_20V_uf"

fil = xw.Book("/Volumes/JACK4/"+file_name+".xlsx")


# sht = file.sheets[0]
# 引用單元格
# rng = sht.range('a1')
# rng = sht['a1']
# rng = sht[0,0] 第一行的第一列即a1,相當於pandas的切片
# 引用區域
# rng = sht.range('a1:a5')
# rng = sht['a1:a5']
# rng = sht[:5,0]

nrows = fil.sheets[2]["a1"].expand("table").rows.count - 2

data = fil.sheets["data"]
side_wb = np.array(data[f""+chr(96+1)+"3:"+chr(96+2)+str(nrows+2)].value)  # wingbase
side_wt = np.array(data[f""+chr(96+3)+"3:"+chr(96+4)+str(nrows+2)].value)  # wingtip
side_te = np.array(data[f""+chr(96+5)+"3:"+chr(96+6)+str(nrows+2)].value)  # hindwing trailing edge
side_tail = np.array(data[f""+chr(96+7)+"3:"+chr(96+8)+str(nrows+2)].value)

top_wb = np.array(data[f""+chr(96+9)+"3:"+chr(96+10)+str(nrows+2)].value)  # wingbase
top_wt = np.array(data[f""+chr(96+11)+"3:"+chr(96+12)+str(nrows+2)].value)  # wingtip
top_te = np.array(data[f""+chr(96+13)+"3:"+chr(96+14)+str(nrows+2)].value)  # hindwing trailing edge
top_tail = np.array(data[f""+chr(96+15)+"3:"+chr(96+16)+str(nrows+2)].value)
# print("finnish colecting data and multiply top with 2180/800")
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
    for j in range(2): #以第一個wingbase為坐標原點
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

    #整合top & side
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


body_vector = wb - tail #body vector
le_vector = wt - wb #vector wing-base to wing-tip 
hind_te_vector = te - wb #vector trailing edge to wing-base

sweeping_angle = []
abdomen_angle = []
flapping_angle = []


for i in range(nrows):
    #calculate sweeping angle
    wingplane_normal_vector = np.cross(le_vector[i], hind_te_vector[i]) #翅膀面法向量
    sw_base_vector = np.cross(wingplane_normal_vector, body_vector[i]) #翅膀面法向量外積身體向量
    len_sw_base = LA.norm(sw_base_vector) #lenth of 翅膀面法向量外積身體向量
    len_le = LA.norm(le_vector[i]) #lenth of vector wing-base to wing-tip
    sw_temp = np.arccos(np.dot(sw_base_vector, le_vector[i]) / (len_sw_base * len_le)) * 180 / np.pi #算角度
    sweeping_angle.append(sw_temp)

    
    #calculate abdomen angle
    abdomen_angle.append(np.arctan(body_vector[i][1] / body_vector[i][0]) * 180 / np.pi)


    #calculate flapping angle
        # calculate vector
    len_body_vector = np.sqrt(body_vector[i][0] ** 2 + body_vector[i][2] ** 2)
    body_right_temp_x = body_vector[i][2] * len_body_vector
    body_right_temp_z = -body_vector[i][0] * len_body_vector
    # body_right_vector = np.array([body_right_temp_x,wb[i][1],body_right_temp_z])
    body_right_vector = np.array([body_right_temp_x, 0, body_right_temp_z])
    len_body_right = np.sqrt(np.dot(body_right_vector, body_right_vector))

    #flap_cos = np.dot(sw_base_vector, body_right_vector) / (len_sw_base * len_body_right)
    flap_sin = LA.norm(np.cross(sw_base_vector, body_right_vector)) / (len_sw_base * len_body_right)
    #flap_temp = np.dot(sw_base_vector, [0, 0, -1]) / (len_sw_base)

    if sw_base_vector[1] > 0:
        flap_temp = np.arccos(flap_cos) * 180 / np.pi
    else:
        flap_temp = -np.arccos(flap_cos) * 180 / np.pi

    flapping_angle.append(-flap_temp)

print("2/2")

#create angle sheet
try :
    print(fil.sheets["angle"])
except :
    fil.sheets.add(name="angle",after=-1)
    print(fil.sheets["angle"])
angle = fil.sheets["angle"]
#create number
for i in range(1,nrows+1):
    angle["a"+str(i+1)].value = str(i)
print("side bar number create finnish")
#flapping_angle 
angle["b"+str(1)].value = "flapping angle"
for i in range(1,nrows+1):
    angle["b"+str(i+1)].value = str(flapping_angle[i-1])
print("flapping angle sheet finnished")
# abdomen_angle 
angle["c"+str(1)].value = "abdomen angle"
for i in range(1,nrows+1):
    angle["c"+str(i+1)].value = str(abdomen_angle[i-1])
print("abdomen angle sheet finnished")
# sweeping_angle 
angle["d"+str(1)].value = "sweeping angle"
for i in range(1,nrows+1):
    angle["d"+str(i+1)].value = str(sweeping_angle[i-1])
print("sweeping angle sheet finnished")



temp = np.arange(0, nrows)
plt.plot(temp, abdomen_angle,label = "abdomen")
plt.plot(temp, flapping_angle,label = "flapping")
plt.plot(temp,sweeping_angle,label = "sweeping")
plt.legend()
plt.show()

