import sys
import os
import warnings
import xlsxwriter as xlw
import pylightxl as xl
import glob
import datetime

warnings.filterwarnings(action='ignore', category=UserWarning)

class Production_day():

    def __init__(self, date, process, Standard, Actual, percent_int, reprocessed):
        # on the 12th of Feb 2024 machine hours are split into main plant hours and r-plant hours.
        self.process = process
        self.standard = Standard
        self.actual = Actual
        self.date = date
        self.rates = {# Rates Changed on the new year 24-25. Rates should be calculated on a case by case basis. will need to look into
            "SAWDUST": 0.206,
            "Ext. Milled OSF": 1.0763,
            "Int. Milled OSF": 0.6346,
            "WF": 0.6608,
            "BENTONITE":0.5187,
            "CAUSTIC":0.482,
            "PHOS ACID":1.8555,
            "INTERMACID":0.5488,
            "SPENT ACID":-0.5,
            'MACHINE HOURS':925,
            'Main Plant Hours':1029,
            'R-Plant hours':364
        }
        self.reprocessed = reprocessed
        self.cost = None
        self.std_cost = None
        self.variance = None
        self.var_cost = None
        self.percent_int = percent_int
        self.clean()
        self.get_cost()
        self.get_stdcost()
        self.get_variance()
        self.get_var_cost()
    
    def clean(self):
        tidy_up = {"Main plant Hours":"Main Plant Hours"}
        
        for i in self.standard:
            if i not in self.rates:
                if i in tidy_up:
                    continue
                else:
                    print(f"The consumable label {i} has been used")
                    new = input("What should this really be: ")
                    tidy_up[i] = new
        
        for i in tidy_up:
            if i in self.actual:
                self.actual[tidy_up[i]] = self.actual[i]
                self.standard[tidy_up[i]] = self.actual[i]
                self.actual.pop(i,None)
                self.standard.pop(i,None)
            else:
                continue
    
    def get_cost(self):
        
        Cost = {}
        rates = self.rates
        for i in self.actual:
            try:
                Cost[i] = self.actual[i] * rates[i]
            except TypeError:
                print(f"The entry of {self.actual[i]} in {i} for actual usage has raised a type error")
                self.actual[i] = float(input("Type what this value should be given the error: "))
                Cost[i] = self.actual[i] * rates[i]

        self.cost=Cost
    
    def get_stdcost(self):
        
        Cost = {}
        rates = self.rates
        for i in self.standard:
            try:
                Cost[i] = self.standard[i] * rates[i]
            except TypeError:
                print(f"The entry of {self.standard[i]} in {i} for standard usage has raised a type error")
                self.standard[i] = float(input("Type what this value should be, given the error: "))
                Cost[i] = self.standard[i] * rates[i]
        self.std_cost = Cost

    def get_variance(self):
        Variance = {}
        for i in self.actual:
            try:
                Variance[i] = self.actual[i] - self.standard[i]
            except TypeError:
                print(f"error for {i} with values in actual{self.actual[i]} and in standard:{self.standard[i]}")
                raise

        self.variance = Variance

    def get_var_cost(self):
        Variance = {}
        for i in self.variance:
            Variance[i] = self.variance[i] * self.rates[i]
        self.var_cost = Variance

    def report(self):
        print(self.date.strftime("%Y-%m-%d")+"\n")
        print("Operational cost: £{}".format(sum(self.cost.values())))
        print("Varitation: £{}".format(sum(self.var_cost.values())))
        print("\n")
        print("\t Variance \t Value")
        for i in self.variance:
            print( i + ": " + str(self.variance[i]) + "\t" + str(self.var_cost[i]))

def pop_empty(list_wt_empty_str):
    
    pop_list = []
    for i, j in enumerate(list_wt_empty_str):
        if j == '':
            pop_list.append(i)

    pop_list.reverse()

    for i in pop_list:
        list_wt_empty_str.pop(i)
    
    return list_wt_empty_str

def pull_consumption(path):
    flag = 1
    print(path)
    file = xl.readxl(path)
    try:
        sheet = file.ws("Morning Meeting")
    except UserWarning:
        sheet = file.ws("Morning Meeting Feb24")


    sheet2 = file.ws("Summary")

    percent_int = sheet2.index(row = 29, col = 20)
    
    reprocessed = {"ELCD":0,
                   "C-carbon":0}
    
    def repr_stf(sheet, column):
        colmun_values = sheet.col(col=column)

        start, end = None, None

        for index, value in enumerate(colmun_values):
            if isinstance(value, str) == False:
                continue

            if value.startswith("IG"):
                start = index + 1
            elif value in ["IT", "R Plant"]:
                end = index
        return start, end
    
    repr_dict = {}

    repr_start, repr_stop = repr_stf(sheet2, flag)
    if repr_start == None:
        flag += 1
        repr_start, repr_stop = repr_stf(sheet2, flag)
        print(repr_start, repr_stop)
        keys = sheet2.range(address=f"B{repr_start+1}:G{repr_start+1}")[0]
        data = sheet2.range(address=f"B{repr_start+2}:G{repr_stop}")
    
    else:
        keys = sheet2.range(address=f"A{repr_start+1}:F{repr_start+1}")[0]
        data = sheet2.range(address=f"A{repr_start+2}:F{repr_stop}")

    for row in data:
        repr_dict[row[0]] = {key:row_el for key, row_el in zip(keys[1:], row[1:])}

    for i in repr_dict:
        if any(substring in repr_dict[i]["New grade"].lower() for substring in ["cnr", "cn5"]):
            try:
                reprocessed["ELCD"] += repr_dict[i]["Qty"]
            except TypeError:
                if repr_dict[i]["Qty"] == "":
                    continue
                else:
                    print(f"The entry {repr_dict[i]["Qty"]} cannot be added to the total from reprocessing {repr_dict[i]["Previous grade"]} ")
                    add_to_ELCD = float(input("What Should it be(Type 0 to skip or Ctrl+C to continue) : "))
                    reprocessed["ELCD"] =+ add_to_ELCD
        elif repr_dict[i]["New grade"] == "":
            continue
        else:
            try:
                reprocessed["C-carbon"] += repr_dict[i]["Qty"]
            except TypeError: 
                print(f"The entry {repr_dict[i]["Qty"]} cannot be added to the total from reprocessing {repr_dict[i]["Previous grade"]} ")
                add_to_Ccarbon = float(input("What Should it be(Type 0 to skip or Ctrl+C to continue): "))
                reprocessed["ELCD"] += add_to_Ccarbon
    
            
    keys = []
    std = []
    act = []
    
    date_end = path.find(" ")
    dot = path.rfind(".")

    date = path[date_end-10:date_end]
    process = path[date_end+1:dot]

    date = datetime.datetime.strptime(date,"%d-%m-%Y")

    for i in range(1,14):
        keys.append(sheet.index(row = 25+i, col =2))
        std.append(sheet.index(row = 25+i, col = 5))
        act.append(sheet.index(row = 25+i, col = 8))
        
    keys = pop_empty(keys)
    std = pop_empty(std)
    act = pop_empty(act)
    for i , j in enumerate(keys):

        if isinstance(std[i],str):
            print(f"{keys[i]} standard use is returning the value {std[i]}")
            std[i] = float(input("what should it actually be: "))
        
        if isinstance(act[i],str):
            print(f"{keys[i]} actual usage is returning the value {act[i]}")
            act[i] = float(input("what should this value actually be: "))

    actual = {}
    standard = {}
    
    for i, x in enumerate(keys):
        actual[x] = act[i]
        standard[x] = std[i]
    print(actual)
    print(standard)

    return Production_day( date, process ,standard, actual, percent_int, reprocessed)

def write_to_xl(days, file=os.path.expanduser('~/Documents/prod_pull.xlsx')):
    
    workbook = xlw.Workbook(file)
    cell_format = workbook.add_format({'bold':True})
    cell_format.set_bg_color('#92D050')
    cell_format.set_text_wrap()
    cell_format.set_align("center")
    cell_format.set_align("vcenter")
    prod = workbook.add_worksheet("Costings")
    
    cell_format.set_text_wrap()
    
    prod.set_column(3, 4*len(days[0].actual)+2,15)
    prod.set_column(0,2,12)
    # prod.set_column(1,1,30)

    date_format = workbook.add_format({'num_format':'yyyy-mm-dd'})
    
    prod.write(2, 0, "Date", cell_format)
    prod.write(2, 1, "Campaign", cell_format)
    prod.write(2, 2, "Process", cell_format)
    for i, x in enumerate(days[-1].actual):
        prod.write(0, i*4+3, x, cell_format)
        # prod.write(2, i*5+3, x + " Standard", cell_format)
        prod.write(2, i*4+3, x + " Actual", cell_format)
        prod.write(2, i*4+4, x + " Cost", cell_format)
        prod.write(2, i*4+5, x + " Variance", cell_format)
        prod.write(2, i*4+6, x + " Variance Value", cell_format)
    prod.write(2, len(days[-1].actual)*4+3, "Total Cost", cell_format)
    prod.write(2, len(days[-1].actual)*4+4, "Total Variance", cell_format)
    prod.write(2, len(days[-1].actual)*4+5, "OSF Int-Ext", cell_format)
    prod.write(2, len(days[-1].actual)*4+7, "ELCD reprocessing", cell_format)
    prod.write(2, len(days[-1].actual)*4+6, "C-carbon reprocessing", cell_format)

    for i , d in enumerate(days):
        prod.write(i+3, 0, d.date.date(), date_format)
        prod.write(i+3, 2, d.process,)
        # for j, k in enumerate(d.standard):
        #     prod.write(i+3, j*5+3, d.standard[k])
        for j , k in enumerate(d.actual):
            prod.write(i+3, j*4+3, d.actual[k])
        for j , k in enumerate(d.cost):
            prod.write(i+3, j*4+4,  d.cost[k])
        for j , k in enumerate(d.variance):
            prod.write(i+3, j*4+5,  d.variance[k])
        for j , k in enumerate(d.var_cost):
            prod.write(i+3, j*4+6,  d.var_cost[k])
        prod.write(i+3,len(d.actual)*4+5, d.percent_int)
        prod.write(i+3,len(d.actual)*4+7, d.reprocessed["ELCD"])
        prod.write(i+3,len(d.actual)*4+6, d.reprocessed["C-carbon"])
        prod.write(i+3,len(d.actual)*4+3, sum(d.cost.values()))
        prod.write(i+3, len(d.actual)*4+4, sum(d.var_cost.values()))
    
    workbook.close()


def pull_frm_args():
    """Takes this script's args, writes an xlsx file containing days production info."""
    # This is the final function that will run when the script is called
    # it pulls in arguments from the command line using the sys library and 
    # pipes them into the pull_consumption function and then appends each Production day
    # into into a list which is used in the final 
    # write_to_xl function which creates the xl file. 
    
    paths = sys.argv
    days = []
    for l in paths:
        try:
            if os.path.isfile(l) and l[-4:] == "xlsx" and "~$" not in l:
                days.append(pull_consumption(l))
        
            elif os.path.isdir(l):
                for m in glob.glob("./"+ l + "/*"):
                    if m[-4:] == "xlsx" and "~$" not in m: 
                        days.append(pull_consumption(m))

        except :
            raise
    if len(days) == 1:
        days[0].report()
    
    elif len(days) > 1:
        write_to_xl(days)

if __name__=="__main__":
    pull_frm_args()
    os.system("start EXCEL.EXE C:\\Users\\ACarruther\\Documents\\prod_pull.xlsx")
