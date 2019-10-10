import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt


file_name = "0831_3_2"
file = xw.Book("/Volumes/jack32/motion_analysis/"+file_name+".xlsx")

sht_flapping = file.sheets["flapping_raw"]
sht_pitching = file.sheets["pitching_raw"]
sht_sweeping = file.sheets["sweeping_raw"]

point_number = 20
flapping_angle = []
pitching_angle = []
sweeping_angle = []

ncolumn = sht_flapping['a1'].expand("table").columns.count

for i in range(ncolumn):

    alpha = chr(ord('a')+i)
    nrows = file.sheets["flapping_raw"][alpha +'1'].expand("table").rows.count

    rng_flapping = sht_flapping[:nrows, i].value
    rng_pitching = sht_pitching[:nrows, i].value
    rng_sweeping = sht_sweeping[:nrows, i].value

    interval = (nrows-1) / point_number
    flapping_angle.append([rng_flapping[0]])
    pitching_angle.append([rng_pitching[0]])
    sweeping_angle.append([rng_sweeping[0]])


    for j in range(1,point_number+1):
        index = int(j*interval+0.5)
        flapping_angle[-1].append(rng_flapping[index])
        pitching_angle[-1].append(rng_pitching[index])
        sweeping_angle[-1].append(rng_sweeping[index])


file.sheets.add("flapping1")
file.sheets.add("pitching1")
file.sheets.add("sweeping1")

interval = 1 / point_number

for j in range(point_number+1):

    file.sheets["flapping1"][j, 0].value = interval * j
    file.sheets["pitching1"][j, 0].value = interval * j
    file.sheets["sweeping1"][j, 0].value = interval * j

    for i in range(len(flapping_angle)):

        file.sheets["flapping1"][j, i + 1].value = flapping_angle[i][j]
        file.sheets["pitching1"][j, i + 1].value = pitching_angle[i][j]
        file.sheets["sweeping1"][j, i + 1].value = sweeping_angle[i][j]



