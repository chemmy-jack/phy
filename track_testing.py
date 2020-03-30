import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt
from numpy import linalg as LA
from scipy.ndimage.interpolation import rotate
from mpl_toolkits.mplot3d import Axes3D
import matplotlib.pyplot as plt
import numpy as np

# read data from excel file

file_name = "0831_4_1"
database = xw.Book("/Volumes/JACK4/"+file_name+".xlsx")
nrows = database.sheets[2]["a1"].expand("table").rows.count 
print(nrows-2)
data = database.sheets["data"]
side_wb = np.array(data[f"a3:b"+str(nrows)].value)  # wingbase
side_wt = np.array(data[f"c3:d"+str(nrows)].value)  # wingtip
side_te = np.array(data[f"e3:f"+str(nrows)].value)  # hindwing trailing edge
side_ta = np.array(data[f"g3:h"+str(nrows)].value)  # tail

top_wb = np.array(data[f"i3:j"+str(nrows)].value)  # wingbase
top_wt = np.array(data[f"k3:l"+str(nrows)].value)  # wingtip
top_te = np.array(data[f"m3:n"+str(nrows)].value)  # hindwing trailing edge
top_ta = np.array(data[f"o3:p"+str(nrows)].value)  # tail
print("finnish colecting data")
#multiply top[0] with 1280/800 and top[1] with 800/600(if resolution weren't the same)(top:800*600, side:1280*800)
#print("finnish colecting data and multiply top with 2180/800")

# calculate origin coordinate
nrows -= 2
o_wb = []
o_wt = []
o_te = []
o_ta = []
for i in range(nrows):
    o_wb.append([-(side_wb[i][0] + top_wb[i][0]) / 2, side_wb[i][1], -top_wb[i][1]])
    o_wt.append([-(side_wt[i][0] + top_wt[i][0]) / 2, side_wt[i][1], -top_wt[i][1]])
    o_te.append([-(side_te[i][0] + top_te[i][0]) / 2, side_te[i][1], -top_te[i][1]])
    o_ta.append([-(side_te[i][0] + top_te[i][0]) / 2, side_ta[i][1], -top_ta[i][1]])
print("finnish calculate origin coordinate")

fig = plt.figure()
ax = fig.add_subplot(111, projection='3d')
for i in range(nrows):
	ax.scatter(o_wb[i][0], o_wb[i][1], o_wb[i][2], c='r', marker='o', s=10) 

ax.set_xlabel('X Label')
ax.set_ylabel('Y Label')
ax.set_zlabel('Z Label')

plt.show()
