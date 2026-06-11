import pyodbc
import numpy as np
import pandas as pd
import os
from scipy.optimize import brentq
import datetime

def phosphate_titre(pH):
    """
    This function is based on a triple sigmoid curve fit to the titration of 
    phosphoric acid. This is based on a transcription of data from an image in 
    "Panov, Alexander & Sci,. (2014). PRACTICAL MITOCHONDRIOLOGY Pitfalls and 
    Problems in Studies of mitochondria with a Description of Mitochondrial 
    Functions." 
    """ 

    x = pH
    a1 = 0.98698832
    a2 = 1.0069641
    a3 = 0.93783363
    k1 = 2.36729059
    k2 = 2.29890918
    k3 = 2.59878959
    pk1 = 2.12
    pk2 = 7.21
    pk3 = 12.32  
    return ((a1/(1+np.exp(- k1 * (x - pk1))))+(a2/(1+np.exp(- k2 * (x - pk2))))+(a3/(1+np.exp(-k3*(x - pk3)))))

def inver_phosphate_titre(OHeq):
    """
    This is the inverse function for the phosphate titrate curve. As far as i
    am aware there are no exact representations for this function. It has been 
    split into 3 sections for the mono, die, and tri sodium salts for the curve
    and so has 2 boundry conditions.

    This function is discontinuous and may return larger erros close to the 
    curve boundries. 
    """
    
    x = OHeq
    a1 = 0.98698832
    a2 = 1.0069641
    a3 = 0.93783363
    k1 = 2.36729059
    k2 = 2.29890918
    k3 = 2.59878959
    pk1 = 2.12
    pk2 = 7.21
    pk3 = 12.32

    if OHeq <= 0.99:
        pH = np.log(a1/x - 1)/k1 + pk1
    elif OHeq <=1.99:
        pH = np.log(a2/(x-a1))/k2 + pk2
    else:
        pH = np.log(a3/(x-a1-a2))/k3 +pk3
    return pH

def NaOH_rho_T_surface(Temp, M_frac):
    """
    This function takes Temperature and Mass fraction data and returns the 
    material density of a caustic solution.
    For this function to be valid the following conditions must be true:
    0=<Temp=<100C
    1=<M_frac=<50 
    """
    
    x = Temp
    y = M_frac
    a0 = 1.00301651e+00
    a1 = -1.73653781e-04
    a2 = -2.83201253e-06
    a3 = 2.17028151e-09
    b1 = 1.12379238e-02
    b2 = 1.29232791e-05
    b3 = -4.59799759e-07
    c1 = 8.25474358e-08
    c2 = 6.19366293e-08
    c3 = -1.57528562e-05
    
    return a0 + a1*x + a2*x**2 + a3*x**3 + b1*y + b2*y**2 +b3*y**3 + c1*x*y**2 + c2 *y*x** 2 + c3 *x*y 

def NaOH_conductivity(Cond):
    """
    This function returns a mass fraction of Caustic solution for a given 
    conductivity. This function is only valid if the concentration is below 
    14% Caustic by weight.
    """    
    x = Cond
    a4 = 1.47268E-21
    a3 = -9.22958E-16
    a2 = 2.02522E-10
    a1 = 5.54378E-06
    a0 = 0.071342345

    mass_frac = a4*x**4 +a3*x**3 +a2*x**2 +a1*x + a0
    return mass_frac/100

def C01_conductivity(C01):
    return 160000 * C01/100

def NaOH_kg_mols(kg):
    return kg/39.997

def NaOH_mols_kg(mols):
    return mols*39.997

def Conc_to_conductivity(Conc, Temp):
    def conc_cond_density(C01):
        
        mass_frac = NaOH_conductivity(CO1_conductivity(C01))
        density = NaOH_rho_T_surface(Temp, mass_frac*100)
               
        return Conc - mass_frac * density
    
    conductivity = brentq(conc_cond_density, 2, 96)
    
    return conductivity

def Conc_to_conductivity(Conc, Temp):
    def conc_cond_density(C01):
        
        mass_frac = NaOH_conductivity(C01_conductivity(C01))
        density = NaOH_rho_T_surface(Temp, mass_frac*100)
               
        return Conc - mass_frac * density
    
    conductivity = brentq(conc_cond_density, 1, 96)
    
    return conductivity

def Tag_query(tag, command="AVG",start=None, duration=None, end=None, resolution=5):
    """
    Takes a tag name, command, duration and end point and submits a sql query 
    returning data summarising the tag over this duration.
    """
    if start is None and duration is None and end is None:
        raise ValueError("You need to provie at least 2 points of refrence in time to use this function")
    
    if start is None:
        if duration is None or end is None:
            raise ValueError("If you provide no start you must provide an end or duration")
        start = end - duration

    elif duration is None:
        if end is None:
            raise ValueError("If no duration is provided you must provide an end with a start")
        duration = start - end
    
    elif end is None:
        end = start + duration
    
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=10.43.8.37;PORT=10014")  # 192.168.9.144
    # start = end - timedelta(days = duration)
    end_str = end.strftime("%Y-%m-%d %H:%M:%S")
    start_str=start.strftime("%Y-%m-%d %H:%M:%S")

    # time period depends on the reporting from the instruments 
    if command not in ('AVG','STDDEV', 'VARIANCE', 'COUNT','SUM', 'MIN', 'MAX'):
        raise ValueError("""This function will only accept the following commands: 
                         COUNT, MIN, MAX, SUM, AVG, STDDEV, VARIANCE""")
    
    sql = "select %s(\"IP_TREND_VALUE\") as \"VALUE\" from \"%s\" "\
            " where \"IP_TREND_TIME\" between TIMESTAMP'%s' and TIMESTAMP'%s'" %(command, tag, start_str, end_str)

    # data = pd.read_sql(sql,con) # Pandas DataFrame with your data!
    
    cursor = con.cursor()
    result = cursor.execute(sql)
    rows = result.fetchall()
    # print(rows)
    cols = []
    for i in result.description:
        cols.append(i[0])
    
    if len(rows) > 0:
        data = pd.DataFrame(data=np.array(rows), columns=cols)
        data['VALUE'] = data['VALUE'].astype(float)
        data['VALUE'] = data['VALUE'].round(4)
        data.name = tag
    
    elif len(rows)==0:
        print("Warning: no changes detected in {} from {} to {} collecting interpolated data".format(tag, start_str, end_str))
    
    # Pandas DataFrame with your data to 4 decimal places!
    return data

def c01_query(tags, start=None, duration=None, end=None):

    if start is None and duration is None and end is None:
        raise ValueError("You need to provie at least 2 points of refrence in time to use this function")
    
    if start is None:
        if duration is None or end is None:
            raise ValueError("If you provide no start you must provide an end or duration")
        start = end - duration

    elif duration is None:
        if end is None:
            raise ValueError("If no duration is provided you must provide an end with a start")
        duration = start - end
    
    elif end is None:
        end = start + duration

    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=10.43.8.37;PORT=10014" )# 192.168.9.144
    end_str = end.strftime("%Y-%m-%d %H:%M:%S")
    start_str=start.strftime("%Y-%m-%d %H:%M:%S")
    
    sql = 'select a.TS "Datetime", a.VALUE "C01", b.VALUE "T40", c.VALUE "F30", d.VALUE "F31"'\
        f'from (select "TS", "VALUE" from history where NAME=\'{tags["C01"]}\' and '\
        f'PERIOD=10 and TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\') a, '\
        f'(select "TS", "VALUE" from history where NAME=\'{tags["T40"]}\' ' \
        f'and PERIOD=10 and TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\') b, '\
        f'(select "TS", "VALUE" from history where NAME=\'{tags["F30"]}\' and '\
        f'PERIOD=10 and TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\') c, '\
        f'(select "TS", "VALUE" from history where NAME=\'{tags["F31"]}\' and '\
        f'PERIOD=10 and TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\') d '\
        'where a.TS = b.TS and b.TS = c.TS and c.TS = d.TS '\
        f'and a.TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\' '\
        f'and b.TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\' '\
        f'and c.TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\' '\
        f'and d.TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\' '\
        'ORDER BY a.TS'
    # print(sql)
    cursor = con.cursor()
    result = cursor.execute(sql)
    rows = result.fetchall()
    
    cols = []
    for i in result.description:
        cols.append(i[0])

    if len(rows)>0:
        data = pd.DataFrame(data=np.array(rows),columns=cols)
        data["C01"] = data["C01"].astype(float)
        data["T40"] = data["T40"].astype(float)
        data["F30"] = data["F30"].astype(float)
        data["F31"] = data["F31"].astype(float)
        return data
    
    elif len(rows)==0:
        print("Warning no data returned from sql query")
        return None

def Product_timeline():
    pH_result = float(input("Product pH result: "))
    start = datetime.datetime.strptime(input("neutralisation start(YY-mm-dd HH:MM): "),'%y-%m-%d %H:%M')
    end = datetime.datetime.strptime(input("neutralisation end(YY-mm-dd HH:MM): "),'%y-%m-%d %H:%M')

    tags = {
        "C01":"WSH1.TK8.DNAOH.AC.01.PV",
        "T40":"WSH1.TK8.DNAOH.TC.01.PV",
        "F30":"WSH1.TK8.DNAOH.FI.01.PV",
        "F31":"WSH1.TK8.DNAOH.FI.02.PV"
    }

    data = c01_query(tags, start=start, end=end)
    print(data["C01"])

    data["Conductivity"] = data.apply(lambda x: C01_conductivity(x.C01), 1 )
    data["M_NaOH"] = data.apply(lambda x: NaOH_conductivity(x.Conductivity),1)
    data["rho_NaOH"] = data.apply(lambda x: NaOH_rho_T_surface(x.M_NaOH, x.T40),1)
    data["Mdot_NaOH"] = data.apply(lambda x: (x.F30 + x.F31) * x.rho_NaOH * x.M_NaOH, 1)
    data["Ndot_NaOH"] = data.apply(lambda x: NaOH_mols_kg(x.Mdot_NaOH),1)
    
    oh_eq = phosphate_titre(pH_result)
    oh_base = phosphate_titre(2.5)
    

    print(data)

    print("Average Caustic delivered over time period = " + str(round(data["Mdot_NaOH"].mean(), 3)) + " kg/hr")

def main():

    # sample_time = datetime.datetime.strptime(input("time samples was taken (YY-mm-dd HH:MM): "),'%y-%m-%d %H:%M')
    # neutralisation_time = sample_time - datetime.timedelta(hours=2.5)
    # neutralisation_duration = datetime.timedelta(hours=0.5)
    # F3031 = 5100 # L/hr
    # C01 = 20.5 # %
    # T07 = 55 # C
    pH_base = 2.5
    F3031 = 1600 + 3500 #Tag_query("WSH1.TK8.DNAOH.FI.01.PV", "AVG", start=neutralisation_time, duration=neutralisation_duration)["VALUE"][0]
    # F3031 += Tag_query("WSH1.TK8.DNAOH.FI.02.PV", "AVG", start=neutralisation_time, duration=neutralisation_duration)["VALUE"][0]
    C01 = 95 #Tag_query("WSH1.TK8.DNAOH.AC.01.PV", "AVG", start=neutralisation_time, duration=neutralisation_duration)["VALUE"][0]
    T40 = 55#Tag_query("WSH1.TK8.DNAOH.TC.01.PV", "AVG", start=neutralisation_time, duration=neutralisation_duration)["VALUE"][0]
    
    print("F31 and F30 total = " + str(F3031)+ " l/hr")
    print("Current C01 = " + str(C01) + " %")
    print("Current T07 = " + str(T40) + " C")


    pH_result = float(input("pH result from L501 : ")) # pH
    pH_target = float(input("pH target : ")) # pH

    NaOH_Mass_frac = NaOH_conductivity(C01_conductivity(C01))
    rho = NaOH_rho_T_surface(T40, NaOH_Mass_frac)
    Caustic_mass_flow = NaOH_Mass_frac * rho * F3031
    Caustic_molar_flow = NaOH_kg_mols(Caustic_mass_flow)

    OHeq_current = phosphate_titre(pH_result)
    OHeq_base = phosphate_titre(pH_base)
    OHeq_target = phosphate_titre(pH_target)

    OHeq_naoh_current = OHeq_current - OHeq_base
    OHeq_naoh_target = OHeq_target - OHeq_base

    Abstract_Phosphoric = Caustic_molar_flow / OHeq_naoh_current
    target_Caustic_molar_flow = OHeq_naoh_target * Abstract_Phosphoric

    target_Caustic_mass_flow = NaOH_mols_kg(target_Caustic_molar_flow)
    Target_Caustic_conc = target_Caustic_mass_flow / F3031


    Target_Conductivity = Conc_to_conductivity(Target_Caustic_conc, T40)

    print("Current Caustic Mass frac = " +  str(NaOH_Mass_frac) + " kg/kg")
    print("Caustic solution Density = "+ str(rho) + " kg/l")
    print("Current Concentration = " + str(NaOH_Mass_frac*rho) + " kg/l")
    print("Caustic_mass_flow = " + str(Caustic_mass_flow)+ " kg/hr")
    print("Caustic Molar flow = " + str(Caustic_molar_flow) + " mol/hr")
    print("OHeq_current = " + str(OHeq_current) + " mol/mol")
    print("OHeq_target = " + str(OHeq_target) + " mol/mol")
    print("Abstract_Phosphoric = " + str(Abstract_Phosphoric) + " mol/hr" )
    print("Target Caustic molar flow = " + str(target_Caustic_molar_flow) + " mol/hr")
    print("Target Caustic Mass flow = "+ str(target_Caustic_mass_flow) + " kg/hr")
    print("Target Caustic Conc = " + str(Target_Caustic_conc) + " kg/l")
    print("Target Conductivity = " + str(Target_Conductivity) + " %")

print(phosphate_titre(6.92))
print(NaOH_conductivity(Cond=160000))

# print(NaOH_rho_T_surface(25,50))

# if __name__=="__main__":
#     main()
    # Product_timeline()
# def twond_derivative_3rd_root(x):
#     a1 = 0.98698832
#     a2 = 1.0069641
#     a3 = 0.93783363
#     k1 = 2.36729059
#     k2 = 2.29890918
#     k3 = 2.59878959
#     pk1 = 2.12
#     pk2 = 7.21
#     pk3 = 12.32  

#     mono = (a1*k1*np.exp(k1*(x+pk1)))/((np.exp(k1*pk1)+np.exp(k1*x))**2)
#     die = (a2*k2*np.exp(k2*(x+pk2)))/((np.exp(k2*pk2)+np.exp(k2*x))**2)
#     tri = (a3*k3*np.exp(k3*(x+pk3)))/((np.exp(k3*pk3)+np.exp(k3*x))**2)
 
#     return (mono - die - tri)

# def twond_derivative_5th_root(x):
#     a1 = 0.98698832
#     a2 = 1.0069641
#     a3 = 0.93783363
#     k1 = 2.36729059
#     k2 = 2.29890918
#     k3 = 2.59878959
#     pk1 = 2.12
#     pk2 = 7.21
#     pk3 = 12.32  

#     mono = (a1*k1*np.exp(k1*(x+pk1)))/((np.exp(k1*pk1)+np.exp(k1*x))**2)
#     die = (a2*k2*np.exp(k2*(x+pk2)))/((np.exp(k2*pk2)+np.exp(k2*x))**2)
#     tri = (a3*k3*np.exp(k3*(x+pk3)))/((np.exp(k3*pk3)+np.exp(k3*x))**2)
 
#     return (- mono - die + tri)

# threerd_root = brentq(twond_derivative_3rd_root, 4,6)

# fiveth_root = brentq(twond_derivative_5th_root, 9,11)

# def check(x):
#     a1 = 0.98698832
#     a2 = 1.0069641
#     a3 = 0.93783363
#     k1 = 2.36729059
#     k2 = 2.29890918
#     k3 = 2.59878959
#     pk1 = 2.12
#     pk2 = 7.21
#     pk3 = 12.32
#     mono = (a1 * k1**2 * np.exp(k1 * (2 * pk1 + x)) - np.exp(k1 * (pk1 + 2 * x)))/(np.exp(k1 * pk1) + np.exp(k1 * x))**3
#     die = (a2 * k2**2 * np.exp(k2 * (2 * pk2 + x)) - np.exp(k2 * (pk2 + 2 * x)))/(np.exp(k2 * pk2) + np.exp(k2 * x))**3
#     tri = (a3 * k3**2 * np.exp(k3 * (2 * pk3 + x)) - np.exp(k3 * (pk3 + 2 * x)))/(np.exp(k3 * pk3) + np.exp(k3 * x))**3
#     return mono + die + tri

# third_check = brentq(check, 4,6)
# fifth_check = brentq(check,9,11)
# print(third_check)
# print(fifth_check)

# print(twond_derivative_1st_root(0))
# print(twond_derivative_1st_root(1))
# print(twond_derivative_1st_root(2))
# print(twond_derivative_1st_root(3))
# print(twond_derivative_1st_root(4))
# print(twond_derivative_1st_root(5))
# print(twond_derivative_1st_root(6))
# print(twond_derivative_1st_root(7))
# print(twond_derivative_1st_root(8))

# print("The first boundry for the inverse function corresponds to pH " + str(threerd_root) + ", [OH]eq " + str(phosphate_titre(threerd_root)))
# print ("The second boundry for the inverse function corresponds to pH " + str(fiveth_root) + ", [OH]eq " + str(phosphate_titre(fiveth_root)))