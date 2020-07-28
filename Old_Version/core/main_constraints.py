import numpy as np
from scipy.optimize import lsq_linear
import measurments

class DW:
    w0 = 136.14
    dhwl = 36 / (w0 + 36)
    hhwl = 9 / (w0 + 9)
    hhwg = 1.5 * 18 / (w0 + 9)
    a3wq1 = 9 / w0
    a3wq2 = 36 / w0


def GPA(wl_id, dwl, dwa, dwh):
    A = np.zeros((4, 4))
    b = np.zeros(4)
    if dwa >= 0:
        A[0, 0] = 0
        A[0, 1] = 0
        A[0, 2] = DW.a3wq1
        A[0, 3] = 0
        A[1, 0] = 0
        A[1, 1] = DW.hhwg
        A[1, 2] = DW.a3wq2
        A[1, 3] = 0
        A[3, 0] = 1
        A[3, 1] = 1
        A[3, 2] = 1
        A[3, 3] = 1
        if wl_id == 1:
            A[2, 0] = DW.dhwl
            A[2 ,1] = DW.hhwl
            A[2, 2] = 0
            A[2, 3] = 0
            b[2] = dwl
        if wl_id == 2:
            A[2, 0] = DW.dhwl
            A[2, 1] = DW.hhwl
            A[2, 2] = (1+DW.a3wq1)*DW.hhwl-DW.a3wq1*dwl/100
            A[2, 3] = 0
            b[2] = dwl
        if wl_id == 3:
            A[2, 0] = DW.dhwl
            A[2, 1] = (1+DW.hhwg)*DW.dhwl-DW.hhwg*dwl/100
            A[2, 2] = (1+DW.a3wq2)*DW.dhwl-DW.a3wq2*dwl/100
            A[2, 3] = 0
            b[2] = dwl
        b[0] = dwa
        b[1] = dwh
        b[3] = 100
    else:
        A[0, 0] = 0
        A[0, 1] = 0
        A[0, 2] = -1
        A[0, 3] = 0
        A[1, 0] = 0
        A[1, 1] = DW.hhwg
        A[1, 2] = -1
        A[1, 3] = 0
        A[3, 0] = 1
        A[3, 1] = 1
        A[3, 2] = 1
        A[3, 3] = 1
        if wl_id == 1:
            A[2, 0] = DW.dhwl
            A[2, 1] = DW.hhwl
            A[2, 2] = 1
            A[2, 3] = 0
            b[2] = dwl
        if wl_id == 2:
            A[2, 0] = DW.dhwl
            A[2, 1] = DW.hhwl
            A[2, 2] = 0.01*dwl
            A[2, 3] = 0
            b[2] = dwl
        if wl_id == 3:
            A[2, 0] = DW.dhwl
            A[2, 1] = (1 + DW.hhwg) * DW.dhwl - DW.hhwg * dwl / 100
            A[2, 2] = dwl / 100
            A[2, 3] = 0
            b[2] = dwl
        b[0] = dwa
        b[1] = dwh
        b[3] = 100
    return A, b

def TGA(dwo, dwa, dwh):
    A = np.zeros((4, 4))
    b = np.zeros(4)
    if dwa >= dwo:
        A[0, 0] = 1
        A[0, 1] = 1
        A[0, 2] = 1
        A[0, 3] = 1
        A[1, 0] = DW.dhwl
        A[1, 1] = DW.hhwl
        A[1, 2] = 0
        A[1, 3] = 0
        A[2, 0] = 0
        A[2, 1] = DW.hhwg - 0.01 * DW.hhwg / DW.dhwl * dwh
        A[2, 2] = DW.a3wq2 - 0.01 * DW.a3wq2 / DW.dhwl * dwh
        A[2, 3] = -1
        A[3, 0] = DW.dhwl / DW.hhwl - 1
        A[3, 1] = 0
        A[3, 2] = DW.a3wq1 - 0.01 * DW.a3wq1 / DW.hhwl * dwa
        A[3, 3] = -1
        b[0] = 100
        b[1] = dwo
        b[2] = 1 / DW.dhwl * dwh - 100
        b[3] = 1 / DW.hhwl * dwa - 100
    else:
        A[0, 0] = 1
        A[0, 1] = 1
        A[0, 2] = 1
        A[0, 3] = 1
        A[1, 0] = DW.dhwl
        A[1, 1] = DW.hhwl
        A[1, 2] = 1
        A[1, 3] = 0
        A[2, 0] = 0
        A[2, 1] = DW.hhwg - 0.01 * DW.hhwg / DW.dhwl * dwh
        A[2, 2] = 0.01 * dwh / DW.dhwl - 1
        A[2, 3] = -1
        A[3, 0] = DW.dhwl / DW.hhwl - 1
        A[3, 1] = 0
        A[3, 2] = 0.01 * dwa / DW.hhwl - 1
        A[3, 3] = -1
        b[0] = 100
        b[1] = dwo
        b[2] = 1 / DW.dhwl * dwh - 100
        b[3] = 1 / DW.hhwl * dwa - 100
    return A, b

gpa_m = measurments.GPA()
com_m = measurments.COM()

GPA_PA = []
lbond = np.zeros(4)
ubond = lbond + 100

for i in range(3):
    for wga in gpa_m[0]:
        for wgh in gpa_m[1]:
            for wl in com_m[i]:
                A, b = GPA(i+1, wl, wga, wgh)
                #p = np.linalg.solve(A, b)
                lsd = lsq_linear(A, b, bounds=(lbond, ubond), lsmr_tol='auto', verbose=0)
                p = lsd.x
                p = p.tolist()
                if wga >= 0:
                    p.insert(3, 0)
                else:
                    p.insert(2, 0)
                GPA_PA.append(p)

output = [[b for b in a if b<0] for a in GPA_PA]
GPA_mean = np.squeeze(np.mean(np.array(GPA_PA),axis=0))
GPA_std = np.squeeze(np.std(np.array(GPA_PA),axis=0))

print("GPA")
for i in range(5):
    print(GPA_mean[i], " +- ", GPA_std[i])

TGA_PA = []

for wlo in com_m[0]:
    for wla in com_m[1]:
        for wlh in com_m[2]:
            A, b = TGA(wlo, wla, wlh)
            lsd = lsq_linear(A, b, bounds=(lbond, ubond), lsmr_tol='auto', verbose=0)
            p = lsd.x
            p = p.tolist()
            if wga >= 0:
                p.insert(3, 0)
            else:
                p.insert(2, 0)
            TGA_PA.append(p)

TGA_mean = np.squeeze(np.mean(np.array(TGA_PA),axis=0))
TGA_std = np.squeeze(np.std(np.array(TGA_PA),axis=0))

print("TGA")
for i in range(5):
    print(TGA_mean[i], " +- ", TGA_std[i])