# -*- coding: utf-8 -*-
"""
Created on Mon Mar 14 17:08:05 2016

@author: Sony
"""
import scipy
import os
import numpy
from scipy.integrate import odeint
import win32com.client
thisdir = os.getcwd()
xl=win32com.client.gencache.EnsureDispatch("Excel.Application")
wb = xl.Workbooks.Open(thisdir +"/"+ "project.xlsm")
xl.Visible = True
sheet=wb.Sheets('Sheet1')
W=sheet.Cells(1,2).Value
rho=sheet.Cells(2,2).Value
Cp=sheet.Cells(3,2).Value
V=sheet.Cells(4,2).Value
Ti=sheet.Cells(5,2).Value
Kc=sheet.Cells(6,2).Value
tauI=sheet.Cells(7,2).Value
taud=sheet.Cells(9,2).Value
tauD=sheet.Cells(8,2).Value
taum=sheet.Cells(10,2).Value
Tr=sheet.Cells(11,2).Value
T_int=sheet.Cells(13,2).Value
TO_int=T_int
Tm_int=T_int
errsum_int=0
on=1
off=0
con=sheet.Cells(12,2).Value
t_int=0
t_end=sheet.Cells(14,2).Value
t_list=[]
q_list=[]
#Initial Condition
f_int=[T_int,TO_int,Tm_int,errsum_int]
t=numpy.arange(t_int,t_end,1)
dt = int(t[1]-t[0])
#define equation
def tank(f,t,W,rho,Cp,V,Ti,Kc,tauI,tauD,taum,Tr,dt):
    T,TO,Tm,errsum=f
# define P PI PID
    pro=Kc*(Tr-Tm)
#    print (pro)
    inte=(Kc/tauI)*errsum
#    print (inte)
    der=int(Kc)*int(taud)*numpy.gradient(Tr-Tm,dt)   
#    print (der)    
    if con=='P':
        q=pro
    elif con=='PI':
        q=pro+inte
    elif con=='PID':
        q=pro+inte+der
    #define ode
    dTdt=(W*Cp*(Ti-T)+q)/rho*V*Cp
    dTOdt=(T-TO-(tauD/2)*dTdt)*2/tauD #dead time
    dTmdt=(TO-Tm)/taum
    derrsumdt=Tr-Tm
    #making the list
    t_list.append(t)
    q_list.append(q)    
    return scipy.array([dTdt,dTOdt,dTmdt,derrsumdt])
#solving ode
Tsol=odeint(tank,f_int,t,args=(W,rho,Cp,V,Ti,Kc,tauI,tauD,taum,Tr,dt))
Ts=Tsol[:,0]
TOs=Tsol[:,1]
Tms=Tsol[:,2]
errsums=Tsol[:,3]   
print (Tms)    

 

