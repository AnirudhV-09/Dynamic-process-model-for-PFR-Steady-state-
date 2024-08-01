# -*- coding: utf-8 -*-
"""
Created on Fri Jul  7 12:58:23 2023

@author: Veliy
"""

#Dynamic process model

import numpy as np
from scipy.integrate import odeint
import matplotlib.pyplot as plt
import xlwings as xw
wb=xw.Book('Process development SOP.xlsm')
sheet=wb.sheets("Dynamic Process Model")

#Dynamic process model(Ideal_dz)

def odes(x,z):
    T = sheet.range("B1").value
    Ea = sheet.range("B2").value
    R = sheet.range("B3").value
    ko = sheet.range("B4").value
    k = ko*pow(2.718, (-Ea/(R*T)))
    Ca = sheet.range("B6").value
    Cb = sheet.range("B5").value
    Q = sheet.range("B7").value
    b = sheet.range("B8").value
    h = sheet.range("B9").value
    r = sheet.range("B11").value
    Ac = sheet.range("B10").value
    u = Q/Ac
    m = sheet.range("B12").value
    n = sheet.range("B13").value
    ra = k*pow(Ca,m)*pow(Cb,n)
    Ha = sheet.range("B14").value
    rho = sheet.range("B16").value
    Cp = sheet.range("B15").value
    U = sheet.range("B18").value
    Ah = sheet.range("B19").value
    Tj = sheet.range("B20").value
    Vr = sheet.range("B21").value
    
    Ca=x[0]
    T=x[1]
    
    dCadz = -ra/u             #Steady state assumption
    dTdz = (ra*Ha)/(rho*Cp*u) #- (U*Ah*(T-Tj))/(Vr*rho*Cp*u)
    
    return [dCadz, dTdz]

x0 = [840, 313.16]

zstart=sheet.range("zstart").value
zend=sheet.range("zend").value
zpoints=len(sheet.range("zstart:zend").value)
z = np.linspace(zstart,zend,zpoints)
sheet.range("zstart:zend").value=z.reshape((zpoints,1))

x = odeint(odes, x0, z)

Ca = x[:,0]
CaL = [p/1000 for p in Ca]
T = x[:,1]

sheet.range("Castart").value=Ca.reshape((zpoints,1))
sheet.range("Tstart").value=T.reshape((zpoints,1))

plt.plot(z, CaL, 'r')
#plt.plot(z, T, 'b')

