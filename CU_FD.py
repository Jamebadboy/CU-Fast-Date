# -*- coding: utf-8 -*-
"""
Created on Thu Jan 25 12:50:03 2018

@author: Jamebadboy
"""
import numpy as np
import pandas as pd

def cal(a,b) :
    try :
        score = 0
        for n in range(len(a)) :
            if(n == 0 and a[n] == b[n]) : return 0 #คนเดียวกัน
            if(n == 1 and a[n] != b[n]) : return 0 #เพศที่ต้องการไม่ตรงกัน
            if(n == 2) :
                age = [int(e.strip()) for e in a[n].strip().split("-")]
                if(len(age) == 1 and int(b[n]) != age[0]) : return 0 #อายุไม่ตรงกัน
                elif(len(age) != 1) :
                    if(age[0]>int(b[n]) or age[1]<int(b[n])) : return 0 #ช่วงอายุไม่ตรงกัน
            if(n == 3) :
                height = [int(e.strip()) for e in a[n].strip().split("-")]
                if(len(height) == 1 and int(b[n]) != height[0]) : return 0 #ความสูงไม่ตรงกัน
                elif(len(height) != 1) :
                    if(height[0]>int(b[n]) or height[1]<int(b[n])) : return 0 #ความสูงไม่ตรงกัน
            if(n >= 4) :    
                a[n] = [int(e) for e in a[n].strip().split(",")]
                if int(b[n]) in a[n] :
                    score += 1
        return score
    except : return "ERROR"

#ดึงข้อมูลจาก Excel
Spec = pd.read_excel("./Spec2.xlsx")
Spec = pd.DataFrame.as_matrix(Spec).astype(str)
Self = pd.read_excel("./Self2.xlsx")
Self = pd.DataFrame.as_matrix(Self).astype(str)
Self = np.delete(Self,2,1) #ลบปีเกิดออก

#ใส่ชื่อใน DataFrame
Matching = pd.DataFrame(np.zeros((len(Self)+1,len(Self)+1), dtype = str))
for x in range(len(Self)) :
    Matching[0][x+1] = Self[x][0]
    Matching[x+1][0] = Self[x][0]

#คำนวณคะแนนใส่ใน DataFrame
for i in range(len(Self)) :
    for j in range(len(Self)) :
        Matching[j+1][i+1] += str(cal(Spec[i].tolist(),Self[j].tolist()))
        Matching[j+1][i+1] += "/"
        Matching[j+1][i+1] += str(cal(Spec[j].tolist(),Self[i].tolist()))

#Export DataFrame เป็น Excel
test = pd.DataFrame(Matching)
writer = pd.ExcelWriter("Matching2.xlsx", engine = "xlsxwriter")
test.to_excel(writer, header=None, index=None)    
writer.save()