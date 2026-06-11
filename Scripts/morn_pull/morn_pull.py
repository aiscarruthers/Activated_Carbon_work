import os
import xlsxwriter as xlw
import pylightxl as xl
import json
import pandas as pd
from xlsxpandasformatter import FormattedWorksheet


script_dir = os.path.dirname(os.path.realpath(__file__))

os.chdir(script_dir)

class Control():
    """
    Control objects structure the parameters required to perform control 
    calculations on Aspen Trends  
    """
    def __init__(self, key, string, mean, variance, target, USL, LSL, units):
        """
        Control objects need to be initialised with the following information:
        
        :param self: the object itself
        :param key: the key that indicates where the output of the control 
                    object should go this takes the form of 3-4 characters from 
                    the start of the trend description(F03 or L302 or FROM).
        :param string: this a string of mean +/- 3 sigma
        :param mean: the mean of the parameter being analysed
        :param variance: the variance of the parameter being analysed
        :param target: the control plan target for the parameter
        :param USL: the Upper Specification Limit for the parameter
        :param LSL: the Lower Specification Limit for the parameter
        :param units: the units of the parameter (eg, kg/hr, l/s)
        """
        self.key = key
        self.string = string
        self.mean = float(mean)
        self.var = float(variance)
        self.target = target
        self.usl = USL
        self.lsl = LSL
        self.units = units
        self.string = None
        self.pp = None
        self.ppu = None
        self.ppl = None
        self.ppk = None
        self.control_calc()
        self.stringify()
        self.values = None
        self.calculate()
        self.cell_format = None
        self.cell_format_key = None
        self.cell_colour()

    def control_calc(self):
        """
        Calculates the control parameters from the mean, the variance, and the 
        specification limits. If variance is 0 this will not generate control 
        values.
        """
        if None not in (self.usl, self.lsl):
            try:
                self.pp = (self.usl -self.lsl) / self.var
                self.ppu = (self.usl-self.mean) * 2 / self.var
                self.ppl = (self.mean - self.lsl) * 2 / self.var
                self.ppk = min(self.ppu, self.ppl)
            except ZeroDivisionError:
                print(f"Variance of {self.key} is 0 meaning control parameters cannot be computed")
        else:
            print(f"{self.key} does not have specification limits... skipping this calculation")

    def stringify(self):
        """Generates the string used to present data"""
        self.string = f"{round(self.mean, 2)} ± {round(self.var, 2)} {self.units}"
    
    def calculate(self):
        """Generates a string without units"""
        self.values = f"{round(self.mean,2)} ± {round(self.var, 2)}"

    def cell_colour(self):
        """
        Determines the cell colour that should be used for this parameter 
        based on the control of the parameter. Green if the parameter is 
        controled with enough precision and centered on target. Orange if 
        the control is precise and not centered and Red if the parameter does 
        not have the required precision. 
        """
        if self.pp == None:
            self.cell_format = {"font_color":"black"}
            self.cell_format_key = '{"font_color":"black"}'

        elif self.ppk > 1.3 :
            self.cell_format = {"bg_color":"green", "font_color":"white"}
            self.cell_format_key = '{"bg_color":"green", "font_color":"white"}'
        elif self.pp > 1.3 :
            self.cell_format = {"bg_color":"orange", "font_color":"white"}
            self.cell_format_key = '{"bg_color":"orange", "font_color":"white"}'
        else:
            self.cell_format = {"bg_color":"red","font_color":"white"}
            self.cell_format_key = '{"bg_color":"red","font_color":"white"}'
        
def main():
    """
    This function uses the Control object to take data from a copy of the Aspen 
    Historgram in a spreadsheet called "data_entry.xlsx" and arrange it in an 
    excel spreadsheet mirroring the morning meeting report format.

    It requires user input for the Recipe no, Acid type, Neutralisation 
    target, and production rate that allows the function to select control plan
    settings and specification limits required for the control calculations.
    """
    recipe = str(input("Recipe No?: " ))
    acid_used = str(input("Fresh or int acid,(lower case)?: "))
    neutralisation = str(input("Neutralisation grade?: "))
    percent_rate = int(input("Rate percent(numbers only eg 100)?:"))/100
    
    with open("Recipe_parameters.json") as f:
        recipe_dict = json.load(f)
    
    strings = {}
    
    strings["Mixing"] = ("Sawdust flows\n"
                         "W03 = {W03}\nW02 = {W02}\nW01 = {W01}\n"
                         "\nAcid flows\n"
                         "F06 = {F06}\nF05 = {F05}\nF04 = {F04}\n"
                         "\nCellulose to acid ratios\n"
                         "R03 = {R03}\nR02 = {R02}\nR01 = {R01}\n"
                         "\nWood Moisture\nM01 = {M01}")

    strings["Activation"] = ("Temperatures\n"
                             "C T03 = {T03}\nB T02 = {T02}\nA T01 = {T01}\n"
                             "\nBack Box\n C T19 = {T19}\nB T18 = {T18}\nA T17 = {T17}\n"
                             "\nGases\nC F03 = {F03}\nB F02 = {F02}\nA F01 = {F01}\n"
                             "\n Extraction\n F16 = {F16}")

    strings["Washing"] = ("Digestor flows\n"
                          "F07 = {F07}\nF09 = {F09}\n"
                          "\nWater washes\n"
                          "F25 = {F25}\nF26 = {F26}\nF27 = {F27}\n"
                          "\nAddtnl Water/Caustic washes\n"
                          "F30 = {F30}\nF31 = {F31}\n"
                          u"\nFirst Wash ΔP P06 = {P06}\n"
                          "Addtnl Wash conduct C01 = {C01}\n"
                          "Final conduct C02 = {C02}\n"
                          "First wash lvl L02 = {L02}")

    strings["Acid"] = ("Acid control\n"
                       "Supply density D01 = {D01}\n correlated D01 = {D01_corr}\n"
                       "Return density D02 = {D02}\n correlated D02 = {D02_corr}\n"
                       "\nAcid feed F11 = {F11}\n"
                       "Acid bleed F28 = {F28}\n"
                       "\nML tank lvl L03 = {L03}\n"
                       "Buffer tank lvl L04 = {L04}\n"
                       "\n Acid recovery efficiency\n"
                       "{A_rec} %")

    strings["Dryer"] = ("Bed temperature\n"
                        "T41A = {T41A}\nT41B = {T41B}\n"
                        "Exhaust temperature\nT06 = {T06}\n"
                        "\nFeed rate L302 speed = {L302}\n"
                        "suction P01 = {P01}")

    strings["Cmill"] = ("Carbon Mill\n"
                        "Classifier speed S53 = {CARM}\n"
                        "Classifier Load I52 = {I52}\n"
                        "TK502 feed rate = {S51}\n"
                        "TK503 feed rate = {S52}\n"
                        "TK504 feed rate = {FROM}\n"
                        "\nCmill air to filter F51 = {F51}\n"
                        "Cmill air to mill fan F52 = {F52}\n"
                        "\nPressure before mill P51 = {P51}\n"
                        "Pressure before classifier P52 = {P52}\n"
                        "Pressure before F501 P53 = {P53}\n"
                        "Pressure after F501 P54 = {P54}\n"
                        "F502 Pressure Drop = {F502}")

    file = xl.readxl("data_entry.xlsx")
    sheet = file.ws("data")
    data_logger = {}
    control_data_logger = {}

    def lims_to_tol(mean, UCL, units):
        """
        Control limits calculated by aspen need to be converted to tolerances
        and returned in a format that can be used by other functions either a
        string for legacy functionality or a dictionary for more recent script 
        versions
        
        :param mean: parameter meam
        :param UCL: parameter Upper Control Limit
        :param units: parameter units.
        """

        if isinstance(mean, str):
            return "nonsense"
        tol = UCL - mean
        string = "{} ± {} {}"
        dictionary = {"mean":mean,
                       "3sig": tol}
        return string.format(round(mean,2),round(tol,2), units), dictionary

    def DEMA(x):
        """Polyfit to convert acid density to mass fraction (DEnsity(kg/m3 or g/l) to MAss fraction)"""
        return round(-527.057716 + 1.074723463 * x - 0.000801043 * x**2 + 2.95816E-07 * x**3 + (-4.2184E-11)* x ** 4, 4)/100

    def cell_acid_ratio(data_logger):
        """
        This Calculates the Dry raw material to dry acid ratio, the purpose of 
        this measure is to standardise changes to acid strength and volume. 
        Historic changes to the acid delivered have been made blind without a 
        real understanding of the measures previously used. Other manufactures 
        have been reported to use this measure and it is likely we will find 
        value in it too

        This function returns data as a string to the data logger and probably 
        should be depreciated.
        
        :param data_logger: A dictionary of Control Objects
        """
        def split_param(param_key, data_logger):
            mean, tol = data_logger[param_key].split()[0,2]
            var = (tol/3)**2
            return mean, var
        
        W01, W01_var = split_param("W01", data_logger)
        W02, W02_var = split_param("W02", data_logger)
        W03, W03_var = split_param("W03", data_logger)
        F04, F04_var = split_param("F04", data_logger)
        F05, F05_var = split_param("F05", data_logger)
        F06, F06_var = split_param("F06", data_logger)
        
        # W01 = float(data_logger["W01"].split()[0]) # kg/hr
        # W02 = float(data_logger["W02"].split()[0])
        # W03 = float(data_logger["W03"].split()[0])

        # W01_var = (float(data_logger["W01"].split()[2])/3)**2 #kg/hr
        # W02_var = (float(data_logger["W02"].split()[2])/3)**2
        # W03_var = (float(data_logger["W03"].split()[2])/3)**2

        # F04 = float(data_logger["F04"].split()[0])/1000 # L/hr to m3/hr
        # F05 = float(data_logger["F05"].split()[0])/1000
        # F06 = float(data_logger["F06"].split()[0])/1000
        
        # F04_var = (float(data_logger["F04"].split()[2])/3000)**2 #L/hr to m3/hr
        # F05_var = (float(data_logger["F05"].split()[2])/3000)**2
        # F06_var = (float(data_logger["F06"].split()[2])/3000)**2
        
        D01 = float(data_logger["D01"].split()[0])  # kg/m3
        D01 = D01_corr(D01)
        D01_var = float(data_logger["D01"].split()[2]) * 1.0288
        D01_DEMA_var = (((D01+D01_var) * DEMA(D01+D01_var) - (D01-D01_var) * DEMA(D01 - D01_var))/6)**2

        D = D01 * DEMA(D01)

        M01 = float(data_logger["M01"].split()[0])/100 # % to fraction
        M01_var = (float(data_logger["M01"].split()[2])/300)**2

        def Ratio(F,W,D=D,M=M01):
            try:
                res = round(F*D/(W*(1-M)),3)
            except ZeroDivisionError:
                res = 0
            return res
        
        R01 = Ratio(F04, W01)
        R02 = Ratio(F05, W02)
        R03 = Ratio(F06, W03)

        # define partial derivatives for variance calculation
        def RparF(W, D=D,M=M01):
            try:
                res = (D/(W*(1-M))) ** 2
            except ZeroDivisionError:
                res = 0
            return res
        
        def RparD(F, W, M=M01):
            try:
                res = (F/(W*(1-M))) ** 2 
            except ZeroDivisionError:
                res = 0
            return res
        
        def RparW(F,W, D=D,M=M01):
            try:
                res = (F*D/((1-M) * (W**2))) ** 2
            except ZeroDivisionError:
                res = 0
            return res
        
        def RparM(F, W, D=D,M=M01):
            try:
                res = (F*D/(W*((1-M)**2))) ** 2
            except ZeroDivisionError:
                res = 0
            return res

        R01_Var = RparF(W01) * F04_var + RparD(F04,W01) * D01_DEMA_var + RparW(F04,W01) * W01_var + RparM(F04,W01) * M01_var
        R02_Var = RparF(W02) * F05_var + RparD(F05,W02) * D01_DEMA_var + RparW(F05,W02) * W02_var + RparM(F05,W02) * M01_var
        R03_Var = RparF(W03) * F06_var + RparD(F06,W03) * D01_DEMA_var + RparW(F06,W03) * W03_var + RparM(F06,W03) * M01_var

        R01_range = round((R01_Var ** 0.5) * 3,3)
        R02_range = round((R02_Var ** 0.5) * 3,3)
        R03_range = round((R03_Var ** 0.5) * 3,3)
        
        data_logger["R01"] = f"{R01} ± {R01_range} ratio"
        data_logger["R02"] = f"{R02} ± {R02_range} ratio"
        data_logger["R03"] = f"{R03} ± {R03_range} ratio"
        return data_logger

    def acid_ratio(control_data_logger):
        W01 = control_data_logger["W01"].mean # kg/hr
        W02 = control_data_logger["W02"].mean
        W03 = control_data_logger["W03"].mean

        W01_var = (control_data_logger["W01"].var / 3)**2 #kg/hr
        W02_var = (control_data_logger["W02"].var / 3)**2
        W03_var = (control_data_logger["W03"].var / 3)**2

        F04 = control_data_logger["F04"].mean /1000 # L/hr to m3/hr
        F05 = control_data_logger["F05"].mean /1000
        F06 = control_data_logger["F06"].mean /1000
        
        F04_var = (control_data_logger["F04"].var /3000)**2 #L/hr to m3/hr
        F05_var = (control_data_logger["F05"].var /3000)**2
        F06_var = (control_data_logger["F06"].var /3000)**2
        
        D01 = control_data_logger["D01"].mean  # kg/m3
        D01 = D01_corr(D01)
        D01_var = control_data_logger["D01"].var * 1.0288
        D01_DEMA_var = (((D01+D01_var) * DEMA(D01+D01_var) - (D01-D01_var) * DEMA(D01 - D01_var))/6)**2

        D = D01 * DEMA(D01)

        M01 = control_data_logger["M01"].mean / 100 # % to fraction
        M01_var = (control_data_logger["M01"].var /300)**2

        def Ratio(F,W,D=D,M=M01):
            try:
                res = round(F*D/(W*(1-M)),3)
            except ZeroDivisionError:
                res = 0
            return res
        
        R01 = Ratio(F04, W01)
        R02 = Ratio(F05, W02)
        R03 = Ratio(F06, W03)

        # define partial derivatives for variance calculation
        def RparF(W, D=D,M=M01):
            try:
                res = (D/(W*(1-M))) ** 2
            except ZeroDivisionError:
                res = 0
            return res
        
        def RparD(F, W, M=M01):
            try:
                res = (F/(W*(1-M))) ** 2 
            except ZeroDivisionError:
                res = 0
            return res
        
        def RparW(F,W, D=D,M=M01):
            try:
                res = (F*D/((1-M) * (W**2))) ** 2
            except ZeroDivisionError:
                res = 0
            return res
        
        def RparM(F, W, D=D,M=M01):
            try:
                res = (F*D/(W*((1-M)**2))) ** 2
            except ZeroDivisionError:
                res = 0
            return res

        R01_Var = RparF(W01) * F04_var + RparD(F04,W01) * D01_DEMA_var + RparW(F04,W01) * W01_var + RparM(F04,W01) * M01_var
        R02_Var = RparF(W02) * F05_var + RparD(F05,W02) * D01_DEMA_var + RparW(F05,W02) * W02_var + RparM(F05,W02) * M01_var
        R03_Var = RparF(W03) * F06_var + RparD(F06,W03) * D01_DEMA_var + RparW(F06,W03) * W03_var + RparM(F06,W03) * M01_var

        R01_range = round((R01_Var ** 0.5) * 3,3)
        R02_range = round((R02_Var ** 0.5) * 3,3)
        R03_range = round((R03_Var ** 0.5) * 3,3)
        
        control_data_logger["R01"] = Control("R01", None, R01, R01_range, None, None, None, "ratio")
        control_data_logger["R02"] = Control("R02", None, R02, R02_range, None, None, None, "ratio")
        control_data_logger["R03"] = Control("R03", None, R03, R03_range, None, None, None, "ratio")
        return control_data_logger
    
    def total_feed_rate(data_logger):
        W01 = float(data_logger["W01"].split()[0]) # kg/hr
        W02 = float(data_logger["W02"].split()[0])
        W03 = float(data_logger["W03"].split()[0])

        W01_var = (float(data_logger["W01"].split()[2])/3)**2 #kg/hr
        W02_var = (float(data_logger["W02"].split()[2])/3)**2
        W03_var = (float(data_logger["W03"].split()[2])/3)**2
        
        Total_feed = W01 + W02 + W03
        Total_feed_var = W01_var + W02_var + W03_var

        data_logger["Total_feed"] = f"{round(Total_feed,2)} ± {round(Total_feed_var, 2)}"
        return data_logger
    
    def plant_feed_rate(control_data_logger):
        W01 = control_data_logger["W01"].mean # kg/hr
        W02 = control_data_logger["W02"].mean
        W03 = control_data_logger["W03"].mean

        W01_var = control_data_logger["W01"].var  #kg/hr
        W02_var = control_data_logger["W02"].var
        W03_var = control_data_logger["W03"].var

        Total_feed = W01 + W02 + W03
        Total_feed_var = W01_var + W02_var + W03_var
        key = "Total_feed"
        string = f"{round(Total_feed,2)} ± {round(Total_feed_var,2)}"
        control_data_logger["Total_feed"] = Control(key, string, Total_feed, Total_feed_var, target=None , USL=None, LSL=None, units="kg/hr")
        return control_data_logger
    
    def D02_corr(D02): # Correlation will need updated as time goes on. Data from 2023 
        return 0.8184 * D02 + 199.88

    def D01_corr(D01):
        return 1.0288 * D01 - 59.227

    def acid_recovery_rep(data_logger):
        D02 = float(data_logger["D02"].split()[0]) # kg/m3
        D02 = D02_corr(D02)

        print(f"{round(D01_corr(1480),2)} ± {round((1.0288 * 5),2)}")

        D02_var = (float(data_logger["D02"].split()[2])*0.8184/3)**2
        
        D2 = D02 * DEMA(D02)
        D2_var = (((D02+D02_var) * DEMA(D02+D02_var) - (D02-D02_var) * DEMA(D02 - D02_var))/6)**2

        F09 = float(data_logger["F09"].split()[0])/1000 # l/hr
        F09_var = (float(data_logger["F09"].split()[2])/3000)**2 #m3/hr

        D01 = float(data_logger["D01"].split()[0]) #kg/m3
        D01 = D01_corr(D01)
        D01_var = (float(data_logger["D01"].split()[2])*1.0288/3)**2

        D1 = D01 * DEMA(D01)
        D1_var = (((D01+D01_var) * DEMA(D01+D01_var) - (D01-D01_var) * DEMA(D01 - D01_var))*1.0288/6)**2

        F04 = float(data_logger["F04"].split()[0])/1000 # m3/hr
        F05 = float(data_logger["F05"].split()[0])/1000 
        F06 = float(data_logger["F06"].split()[0])/1000 
        
        F04_var = (float(data_logger["F04"].split()[2])/3000)**2 #m3/hr
        F05_var = (float(data_logger["F05"].split()[2])/3000)**2
        F06_var = (float(data_logger["F06"].split()[2])/3000)**2

        AU01 = F04 + F05 + F06
        AU01_var = F04_var + F05_var + F06_var
        print(AU01_var)
        PO4_in = D1 * AU01
        PO4_in_var = (AU01 ** 2) * D1_var + (D1 ** 2) * AU01_var
        print(f"PO4 in = {round(PO4_in,2)} ± {round(PO4_in_var ** 0.5 * 3,2)}")
        PO4_out = D2 * F09
        PO4_out_var = (F09 ** 2) * D2_var + (D2 ** 2) * F09_var
        print(f"PO4 out = {round(PO4_out,2)} ± {round(PO4_out_var ** 0.5 * 3,2)}")

        Rec = round( PO4_out , 2)
        Rec_var = round(PO4_out_var, 2)

        Rec_eff = round(100 * PO4_out/PO4_in, 2) # fraction to Percentage
        Rec_eff_var = (PO4_in ** -2) * PO4_out_var + ((PO4_out/(PO4_in**2))**2) * PO4_in

        Rec_eff_range = round(100 * ((Rec_eff_var ** 0.5) * 3),2) # fraction to percentage

        data_logger["Acid_recovery_rate"] = f"{Rec} ± {Rec_var}"
        data_logger["A_rec"] = f"{Rec_eff} ± {Rec_eff_range}"
        data_logger["D01_corr"] = f"{round(D01,2)} ± {round(D01_var,2)} kg/m3"
        data_logger["D02_corr"] = f"{round(D02,2)} ± {round(D02_var,2)} kg/m3"

        return data_logger        

    def control_acid_recovery_rep(control_data_logger):
        D02 = control_data_logger["D02"].mean # kg/m3
        D02 = D02_corr(D02)

        print(f"{round(D01_corr(1480),2)} ± {round((1.0288 * 5),2)}")

        D02_var = (control_data_logger["D02"].var * 0.8184 / 3 ) ** 2
        
        D2 = D02 * DEMA(D02)
        
        D2_var = (((D02+D02_var) * DEMA(D02+D02_var) - (D02-D02_var) * DEMA(D02 - D02_var))/6)**2

        F09 = control_data_logger["F09"].mean / 1000 # m3/hr
        F09_var = (control_data_logger["F09"].var / 3000) ** 2 #m3/hr

        D01 = control_data_logger["D01"].mean #kg/m3
        D01 = D01_corr(D01)
        D01_var = ((control_data_logger["D01"].var)*1.0288/3)**2

        D1 = D01 * DEMA(D01)
        D1_var = (((D01+D01_var) * DEMA(D01+D01_var) - (D01-D01_var) * DEMA(D01 - D01_var))*1.0288/6)**2

        F04 = control_data_logger["F04"].mean / 1000 # m3/hr
        F05 = control_data_logger["F05"].mean / 1000 
        F06 = control_data_logger["F06"].mean / 1000 
        
        F04_var = (control_data_logger["F04"].var / 3000) ** 2 #m3/hr
        F05_var = (control_data_logger["F05"].var / 3000) ** 2
        F06_var = (control_data_logger["F06"].var / 3000) ** 2

        AU01 = F04 + F05 + F06
        AU01_var = F04_var + F05_var + F06_var
        # print(AU01_var)
        PO4_in = D1 * AU01

        PO4_in_var = (AU01 ** 2) * D1_var + (D1 ** 2) * AU01_var

        print(f"PO4 in = {round(PO4_in,2)} ± {round(PO4_in_var ** 0.5 * 3,2)}")
        PO4_out = D2 * F09
        PO4_out_var = (F09 ** 2) * D2_var + (D2 ** 2) * F09_var
        print(f"PO4 out = {round(PO4_out,2)} ± {round(PO4_out_var ** 0.5 * 3,2)}")

        Rec_eff = round(100 * PO4_out/PO4_in, 2) # fraction to Percentage
        Rec_eff_var = (PO4_in ** -2) * PO4_out_var + ((PO4_out/(PO4_in**2))**2) * PO4_in

        Rec_eff_range = round(100 * ((Rec_eff_var ** 0.5) * 3),2) # fraction to percentage

        control_data_logger["A_rec"] = Control("A_rec", None, Rec_eff, Rec_eff_range, None, None, None, "%")
        control_data_logger["D01_corr"] = Control("D01_corr", None, D01, D01_var,
                                                recipe_dict[recipe]["D01_corr"]["x"],
                                                recipe_dict[recipe]["D01_corr"]["USL"],
                                                recipe_dict[recipe]["D01_corr"]["LSL"],
                                                "kg/m3")
        control_data_logger["D02_corr"] = Control("D02_corr", None, D02, D02_var, None, None, None, "kg/m3")

        return control_data_logger        
    
    # extract data to dictionary    
    print("\n")

    for row in sheet.rows:
        if row[0] == "":
            continue
        mean = row[3]
        units = row[5]
        ucl = row[2]

        print( "start"+ row[1] + " " + str(lims_to_tol(mean, ucl,units)))

        info_str, info_dict = lims_to_tol(mean,ucl,units)

        # Now we generate a key from the first 4 characters of the tag, removing
        # any white space, this is fine for most tags but the classifier speed 
        # is formatted differently as well as the V601 scrubber and demister 
        # pressure drops. The first 4 characters of these 2 tags are identicle 
        # and so an exception must be made for them. 
         
        key = row[1][:4]
        key = key.replace(" ","")
        if key == "V601":
            if len(row[1])<23:
                key = "V601"
            else:
                key = "V601 Demist"
        data_logger[key] = info_str

        quantitative_parameters = ("W01","W02", "W03", "F04","F05","F06","F03",
                                   "F02","F01","F11","F28","F25","F26","F27",
                                   "F09","F07","F08","L302","F31","F30",
                                   "L302")
        
        if key in recipe_dict[recipe]:
            if key == "C01":
                target = recipe_dict[recipe][key][neutralisation]["x"]
                lsl = recipe_dict[recipe][key][neutralisation]["LSL"]
                usl = recipe_dict[recipe][key][neutralisation]["USL"]
            elif key in ("F11","F28"):
                target = recipe_dict[recipe][key][acid_used]["x"]
                lsl = recipe_dict[recipe][key][acid_used]["LSL"]
                usl = recipe_dict[recipe][key][acid_used]["USL"]

            else:
                try:
                    target = recipe_dict[recipe][key]["x"]
                except KeyError:
                    print(f"Error accessing recipe dict where key = {key}")
                try:
                    usl = recipe_dict[recipe][key]["USL"]
                    lsl = recipe_dict[recipe][key]["LSL"]
                except KeyError:
                    print(f"No Specification limits found for {key} in recipe {recipe}.")
                    usl = None
                    lsl = None
            
            if key in quantitative_parameters:
                try:
                    lsl = lsl*percent_rate
                    usl = usl*percent_rate
                except TypeError:
                    pass
        
        else:
            target = None
            lsl = None
            usl = None

        control_data_logger[key] = Control(key,info_str, mean, 
                                           info_dict["3sig"], target, usl,
                                           lsl, units)

    # Perform Calculations on data
    data_logger = cell_acid_ratio(data_logger)
    data_logger = acid_recovery_rep(data_logger)
    data_logger = total_feed_rate(data_logger)

    control_data_logger = acid_ratio(control_data_logger)
    control_data_logger = control_acid_recovery_rep(control_data_logger)
    control_data_logger = plant_feed_rate(control_data_logger)
    
    # writing data to first report format   
    for i in strings:
        strings[i] = strings[i].format(**data_logger)

    # writing formats to excel sheets
    workbook = xlw.Workbook("morning_report.xlsx")
    report = workbook.add_worksheet("report_pre_2024")
    cell_format = workbook.add_format({'text_wrap':True, 'valign':'center',
                                       'align':'vcenter'})
    report.set_column(0,5,35)
    # writing first format to excel sheet
    for i , j in enumerate(strings):
        report.write(0,i,strings[j],cell_format)

    C_carbons_order = [
                        "Total_feed",
                        "W01","W02","W03", "Space_1", "F04", "F05", "F06", 
                        "Space_2", "Space_3",
                        "R01","R02","R03", "M01", # Mixing
                        "Space_4", "Space_5", "F03", "Space_6","T03","T19",
                        "Space_7", "Space_8", "F02", "Space_9", "T02","T18", 
                        "Space_10", "Space_11",
                        "F01", "Space_12", "T01", "T17","Space_13", # Kilning
                        "F16", "V601", "V601 Demist", "Space_14",
                        "Space_15", "D01", "D01_corr", "F11", "F28", "F09", "F07", 
                        "D02", "D02_corr",  
                        "Space_16", "F25", "F26", "F27", "F08","F30", "F31",
                        "C01","C02", "Space_17", # Washing and Acid control
                        "Space_18", "L302","T06", "T41B","P01", #Drying
                        "Space_19", "Space_20", "S51","S52","FROM", "P51", "F502" # Milling 
                        ]

    C_Carbons_Glossary = {
        "Total_feed":"Plant Feed rate",
        "W01":"W01 Wood to Kiln A","W02":"W02 Wood to Kiln B",
        "W03":"W03 Wood to Kiln C", "Space_1":None, 
        "F04":"F04 Acid to Kiln A", "F05":"F05 Acid to Kiln B", 
        "F06":"F06 Acid to Kiln C", "Space_2":None, 
        "Space_3":"Dry_Sawdust to Dry Acid Ratio",
        "R01":"R01 A Kiln Ratio","R02":"R02 B Kiln Ratio",
        "R03":"R03 C Kiln Ratio", "M01":"M01 Sawdust Moisture", # Mixing
        "Space_4":None, "Space_5":None, "F03":"F03 Gas Flow", # Kiln C 
        "Space_6":"T03 SP","T03":"T03 Activation Temp",
        "T19":"T19 Back box Temp", 
        "Space_7":None, "Space_8":None, "F02":"F02 Gas Flow", # Kiln B
        "Space_9":"T02 SP", "T02":"T02 Activation Temp",
        "T18":"T18 Backbox Temp", 
        "Space_10":None, "Space_11":None, "F01":"F01 Gas Flow",# Kiln A
        "Space_12":"T01 SP", "T01":"T01 Activation Temp",
        "T17":"T17 Backbox Temp","Space_13":None, 
        "F16":"F16 Abatement flow from Kilns", 
        "V601":"Kiln Scrubber Pressure drop", 
        "V601 Demist":"Demister PRessure drop", "Space_14":None, # Abatement 
        "Space_15":"D01 Target", "D01":"D01 Supply Acid Density", 
        "D01_corr":"D01 Estimated actual", 
        "F11":"F11 Acid addition", "F28":"F28 Acid Bleed", 
        "F09":"F09 Recovered Acid", "F07":"F07 Digestor feed", 
        "D02":"D02 Recovered acid", "D02_corr":"D02 Estimated Actual", #Acid Control 
        "Space_16":None, "F25":"F25 Water Wash", 
        "F26":"F26 Water Wash", "F27":"F27 Water Wash", "F08":"F08 Total Water to Prayon",
        "F30":"F30 Additional Wash", "F31":"F31 Additional Wash",
        "C01":"C01 Caustic wash conductivity",
        "C02":"C02 Wash off conductivity", "Space_17":None, # Washing
        "Space_18":None, "L302":"L302 Dryer Feed", 
        "T06":"T06 Exhaust temperature", 
        "T41B":"T41 Bed temperature", "P01":"P01 Dryer Hood Pressure","Space_19":None, # Dryer 
        "Space_20":"Feeding rate", "S51":"S51 TK502 feed rate",
        "S52":"S52 TK503 feed rate","FROM":"S55 LAC feed rate", 
        "P51":"P51 Mill pressure", "F502":"F502 DeltP" #Milling
    }

    Cdf_grouping = {
        "Total_feed":"Mixing",
        "W01":"Mixing", "W02":"Mixing", "W03":"Mixing", "Space_1":"Mixing", 
        "F04":"Mixing", "F05":"Mixing", "F06":"Mixing", "Space_2":"Mixing", 
        "Space_3":"Mixing", "R01":"Mixing", "R02":"Mixing", "R03":"Mixing", 
        "M01":"Mixing","Space_4":"Mixing", #Mixing
        "Space_5":"KilnC", "F03":"KilnC", "Space_6":"KilnC","T03":"KilnC",
        "T19":"KilnC", "Space_7":"KilnC", #Kiln C
        "Space_8":"KilnB", "F02":"KilnB", "Space_9":"KilnB", "T02":"KilnB",
        "T18":"KilnB", "Space_10":"KilnB", # Kiln B 
        "Space_11":"KilnA", "F01":"KilnA", "Space_12":"KilnA", "T01":"KilnA", 
        "T17":"KilnA","Space_13":"KilnA", # Kiln A
        "F16":"Abatement", "V601":"Abatement", "V601 Demist":"Abatement", 
        "Space_14":"Abatement", #Abatement
        "Space_15":"Acid Mgmt", "D01":"Acid Mgmt", "D01_corr":"Acid Mgmt", 
        "F11":"Acid Mgmt", "F28":"Acid Mgmt", "F09":"Acid Mgmt", 
        "F07":"Acid Mgmt", "D02":"Acid Mgmt", "D02_corr":"Acid Mgmt", 
        "Space_16":"Acid Mgmt", # Acid Mgmt
        "F25":"Washing", "F26":"Washing", "F27":"Washing", "F08":"Washing","F30":"Washing", 
        "F31":"Washing","C01":"Washing","C02":"Washing", "Space_17":"Washing", # Washing
        "Space_18":"Drying", "L302":"Drying","T06":"Drying", "T41B":"Drying","P01":"Drying",
        "Space_19":"Drying", #Drying
        "Space_20":"Milling", "S51":"Milling","S52":"Milling","FROM":"Milling",
        "P51":"Milling", "F502":"Milling" # Milling 

    }

    Cdf_index = [C_Carbons_Glossary[i] for i in C_carbons_order]

    Cdf_data = []
    for i in C_carbons_order:
        if "Space" in i:
            Cdf_data.append([None, None, None, None, None,
            None, None, Cdf_grouping[i]])
        else:
            Cdf_data.append(
                [C_Carbons_Glossary[i], control_data_logger[i].values, 
                control_data_logger[i].units, control_data_logger[i].usl, 
                control_data_logger[i].lsl, control_data_logger[i].pp, 
                control_data_logger[i].ppk, Cdf_grouping[i]])

    Cdf_columns = ["Name", "Value","Units", "USL", "LSL", "PP", "PPK", "Grouping",]
    Cdf = pd.DataFrame(Cdf_data,columns=Cdf_columns)

    Cdf["new group"] = Cdf["Grouping"].ne(Cdf["Grouping"].shift(1).bfill()).astype(int)
    
    mid_cells = Cdf.index[Cdf["new group"]== 0].to_list()
    bottom_cells = Cdf.index[Cdf["new group"]== 1].to_list()
    bottom_cells.append(69)
    mid_cells.append(68)    # final cells for the report format

    writer = pd.ExcelWriter("morning_report_v3.xlsx", engine='xlsxwriter')
    Cdf.to_excel(writer, index=True, sheet_name="report")
    wkbk = writer.book
    wksht = writer.sheets["report"]

    mid_format = {"bottom":1,"top":1,"left":2,"right":2}
    bottom_format = {"bottom":2,"top":1,"left":2,"right":2} # 

    mid = wkbk.add_format(mid_format)
    bottom = wkbk.add_format(bottom_format)
    
    for i in mid_cells:
        wksht.set_row(i,None,mid)
    for i in bottom_cells:
        wksht.set_row(i,None,bottom)

    Good = wkbk.add_format({"font_color":"white","bg_color":"green"})
    NotGood = wkbk.add_format({"font_color":"white","bg_color":"orange"})
    Bad = wkbk.add_format({"font_color":"white","bg_color":"red"})

    wksht.conditional_format('C2:C69',{'type':'formula',
                                       'criteria':'=AND(G2>=1.3,H2>=1.3)',
                                       'format':Good})
    wksht.conditional_format('C2:C69',{'type':'formula',
                                       'criteria':'=AND(G2>=1.3,H2<1.3)',
                                       'format':NotGood})
    wksht.conditional_format('C2:C69',{'type':'formula',
                                       'criteria':'=AND(G2<1.3,H2<1.3,H2<>"")',
                                       'format':Bad})

    formattedWorksheet = FormattedWorksheet(wksht, wkbk, Cdf, hasIndex=True)
    formattedWorksheet.format_cols(colFormatList=["Name","Value", "Units"], 
        colPatternFormatList={"border":True})    
    formattedWorksheet.format_add_separation_border_between_groups('Grouping')
    
    writer._save()

    os.system("start EXCEL.EXE morning_report_v3")

if __name__ == "__main__":
    main()