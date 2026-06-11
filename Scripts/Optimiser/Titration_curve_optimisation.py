from scipy.optimize import curve_fit
import numpy as np
import pandas as pd
import pylightxl as xl
import os

# This relocates the working directory to the bin and Script folder
script_dir = os.path.dirname(os.path.realpath(__file__))
os.chdir(script_dir)

# This function defines the expected shape of the data
def triple_sigmoid(x , a1, a2, a3, k1, k2, k3):
    return ((a1/(1+np.exp(- k1 * (x - 2.12))))+(a2/(1+np.exp(- k2 * (x - 7.21))))+(a3/(1+np.exp(-k3*(x- 12.32)))))

def Phosphate_model():
    data = pd.read_csv("PO4_titration_curve.csv")
    ydata = data["Equivalent OH-"] # OH
    xdata = data["pH"] # pH
    p0 = [1,1,1,2,2,2]
    popt, pcov = curve_fit(triple_sigmoid, xdata, ydata, p0=p0, maxfev=10000)
    
    string = f"a1 = {popt[0]}\na2 = {popt[1]}\na3 = {popt[2]}\n"\
        f"k1 = {popt[3]}\nk2 = {popt[4]}\nk3 = {popt[5]}"
    
    print(string)
    return

def triple_exponential(x, a1,a2,a3,b1,b2,b3):
    return(a1 * np.exp( b1 * x ) + a2 * np.exp( b2 * x )+ a3 * np.exp( b3 * x ) )

def forth_order_poly(x, a0, a1, a2, a3, a4):
    return (a0 + a1*x + a2*x**2 + a3*x**3 +a4*x**4)

def NaOH_conductivity_model():
    data = pd.read_csv("NaOH_conductivity_curve_cl4invers.csv")

    ydata = data["Weight%"]
    xdata = data["microS/cm"]

    p0 = [0.07,0,0,0,0]
    popt, pcov = curve_fit(forth_order_poly, xdata, ydata, p0=p0, maxfev=10000)
    
    string = f"a0 = {popt[0]}\na1 = {popt[1]}\na2 = {popt[2]}\n"\
        f"a3 = {popt[3]}\na4 = {popt[4]}"

    print(string)
    return

def naoh_rho_T_surface(X, a0, a1, a2, a3, b1, b2, b3, c1, c2, c3):
    x , y = X
    return a0 + a1*x + a2*x**2 + a3*x**3 + b1*y + b2*y**2 +b3*y**3 + c1*x*y**2 + c2 *y*x** 2 + c3 *x*y 

def naoh_rho_T_model():
    data = pd.read_csv("NaOH_density_data.csv")
    rho = data.iloc[:,1:].to_numpy().flatten()

    temps = list(data.columns[1:])
    temps = [float(x) for x in temps]
    conc = list(data.iloc[:,0])

    X = [(i,j) for j in conc for i in temps]

    temps = [i[0] for i in X]
    conc = [i[1] for i in X]
    p0 = [1,1,1,1,1,1,1,1,1,1]
    popt, pcov = curve_fit(naoh_rho_T_surface, (temps, conc), rho, p0=p0, maxfev=10000)
    
    string = f"a0 = {popt[0]}\na1 = {popt[1]}\na2 = {popt[2]}\n"\
        f"a3 = {popt[3]}\nb1 = {popt[4]}\nb2 = {popt[5]}\nb3 = {popt[6]}\n"\
        f"c1 = {popt[7]}\nc2 = {popt[8]}\nc3 = {popt[9]}"
    print(string)
    return

def sixth_order_poly(x, a0, a1, a2, a3, a4, a5, a6):
    return (a0 + a1*x + a2*x**2 + a3*x**3 + a4*x**4 + a5*x**5 + a6*x**6)

def H3PO4_conductivity_modle():
    """
    Data from the spreadsheet H3PO4_conductivity_curve_cl4invers.csv is only 
    applicable where concentrations are less than 50% if there is reason to 
    belive that the concentration would be above this level then the 
    conductivity will provide a falsly low result. 
    """
    
    data = pd.read_csv("H3PO4_conductivity_curve_cl4invers.csv")

    ydata = data["Weight%"]
    xdata = data["microS/cm"]
    
    p0 = [1,1,1,1,1,1,1]
    popt, pcov = curve_fit(sixth_order_poly, xdata, ydata, p0=p0, maxfev=10000)
    
    string = f"a0 = {popt[0]}\na1 = {popt[1]}\na2 = {popt[2]}\n"\
        f"a3 = {popt[3]}\na4 = {popt[4]}\na5 = {popt[5]}\na6 = {popt[6]}"
    print(string)
    return

H3PO4_conductivity_modle()