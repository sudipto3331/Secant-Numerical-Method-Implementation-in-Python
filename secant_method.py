# -*- coding: utf-8 -*-
"""
Created on Sun Aug 07 19:40:50 2023

@author: Sudipto
"""
#Import libraries as necessary
import math
import numpy as np
from xlwt import Workbook

xl=float(input ('Enter 1st initial value: '))   #1st input
print(xl)
xu=float(input ('Enter 2nd initial value: '))   #2nd input

err=float(input('Enter desired percentage relative error: '))
ite=int(input('Enter number of iterations: '))
#initialization
x_l=np.zeros([ite])
x_u=np.zeros([ite])
x_c=np.zeros([ite])

f_xl=np.zeros([ite])
f_xu=np.zeros([ite])
f_xc=np.zeros([ite])

rel_err=np.zeros([ite])
itern=np.zeros([ite])
x_l[0]=xl
x_u[0]=xu

#begin iteration   
for i in range(ite):
    itern[i]=i+1
    
    f_xl[i]=(667.38/x_l[i])*(1-math.exp(-0.146843*x_l[i]))-40
    f_xu[i]=(667.38/x_u[i])*(1-math.exp(-0.146843*x_u[i]))-40
    
    #Secant Formula
    x_c[i]=x_u[i]-((f_xu[i]*(x_l[i]-x_u[i]))/(f_xl[i]-f_xu[i]))
    f_xc[i]=(667.38/x_c[i])*(1-math.exp(-0.146843*x_c[i]))-40
    #calculating error    
    if i>0:
        rel_err[i]=((x_c[i]-x_c[i-1])/x_c[i])*100
    #terminate if error criteria meets
    if all ([i>0, abs(rel_err[i])<err]):
        break 
    elif f_xc[i]==0:
        break
   
    if i==ite-1:
        break

    x_l[i+1]=x_u[i]
    x_u[i+1]=x_c[i]
        
wb = Workbook()
  

sheet1 = wb.add_sheet('Sheet 1')
num_of_iter=i

sheet1.write(0,3,'Secant')
sheet1.write(0,4,'Method')


sheet1.write(1,0,'Number of iteration')
sheet1.write(1,1,'x_l')
sheet1.write(1,2,'x_u')
sheet1.write(1,3,'x_c')
sheet1.write(1,4,'f(x_l)')
sheet1.write(1,5,'f(x_u)')
sheet1.write(1,6,'f(x_c)')
sheet1.write(1,7,'Relative error')
  
for n in range(num_of_iter+1):
    
    sheet1.write(n+2,0,itern[n])
    sheet1.write(n+2,1,x_l[n])
    sheet1.write(n+2,2,x_u[n])
    sheet1.write(n+2,3,x_c[n])
    sheet1.write(n+2,4,f_xl[n])
    sheet1.write(n+2,5,f_xu[n])
    sheet1.write(n+2,6,f_xc[n])
    sheet1.write(n+2,7,rel_err[n])

sheet1.write(n+4,2,'The')
sheet1.write(n+4,3,'root')
sheet1.write(n+4,4,'is')
sheet1.write(n+4,5,x_c[i])

wb.save('secant.xls')
