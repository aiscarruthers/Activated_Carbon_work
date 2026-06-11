import pylightxl as xl
import os
import re
import datetime as dt
class MainPlantWorkOrder():
    def __init__(self, packed, ID,grade, RawMat, Std_RawMat, Hours, Std_Hours, wo_dt):
        """
        Here we are Trying to define all the important information in a Work order.
        This script was written with C-carbon products in mind and so can only consider
        the raw materials and by-products of Sawdust, Caustic, Fresh, Intermediate
        and spent acid.

        This Object Class is a means to define and agregate data about WOs from 
        Blairs File. Any structural Changes to Blair's file will need to be accomodated
        for in this file.
        """
        self.tonnage = packed
        self.wono = ID
        self.grade = grade
        self.rawmat = RawMat
        self.stdrawmat = Std_RawMat
        self.hours = Hours
        self.stdhours = Std_Hours
        self.datetime = wo_dt
        self.datestr = ""
        self.ratios = None
        self.std_ratios = None
        self.route = None
        self.std_route = None
        self.Ratios()
        self.Routing()
        self.strfydt()


    def Ratios(self):
        try:
            self.ratios = {i:round(float(self.rawmat[i]) / float(self.tonnage), 5) for i in self.rawmat}
            self.std_ratios = {i:round(float(self.stdrawmat[i]) / float(self.tonnage), 5) for i in self.stdrawmat}
        except:
            print("difficulty calculating ratios for this WO presenting collected data:")
            print ("WO tonnage: "+str(self.tonnage))
            for i in self.rawmat:
                print(f"{i} : {self.rawmat[i]} -- {self.stdrawmat[i]}")

    def Routing(self):
        self.route = round(1000 * float(self.hours) / float(self.tonnage), 5)
        self.std_route = round(1000 * float(self.stdhours) / float(self.tonnage), 5)

    def strfydt(self):
        if isinstance(self.datetime, dt.datetime):
            self.datestr = self.datetime.strftime("%Y-%m-%d %H:%M")
    
def insert_after_first_slash(input_string, year):
    parts = input_string.split('/', 1)  # Split the string at the first slash
    # print(parts)
    if "202" in parts[-1]:
        parts[1] = parts[1].replace('202', '2')
    
    if len(parts) > 1 and f"/{year} " not in parts[1]:
        parts[1] = parts[1].replace(' ', f'/{year} ', 1)
    
    parts[1] = re.sub(r'\s+', ' ', parts[1])
    parts = [i.rstrip() for i in parts]          # Insert '/24' after the first number
    print("start " +'/'.join(parts)+ " end")
    return '/'.join(parts)

def cleanup(input_string):
    # print(input_string)
    input_string = re.sub(r'\s+', ' ', input_string)

    return input_string.rstrip()

def std_date(date_string):
    date, time = date_string.split(" ")
    first, second = date.split("/")
    if len(first)==1:
        first = "0"+first
    if len(second) == 1:
        second = "0"+second
    return first + "/" + second + " " + time

def swap_date_time(date_string):
    time, date = date_string.split(" ")
    first, second = date.split("/")
    if len(first)==1:
        first = "0"+first
    if len(second) == 1:
        second = "0"+second
    return first + "/" + second + " " + time

def main():
    db = xl.readxl("C:\\Users\\ACarruther\\OneDrive - NORIT AMERICAS INC\\Documents\\BOMs\\C Carbon closures_230725_v2.xlsx")
    
    WO_list = []
    report = {}
    for sheet in db.ws_names:
        init_wo_length = len(WO_list)
        print(sheet)
        ws = db.ws(ws=sheet)
        for j , flag in enumerate(ws.col(col=20)):
            if flag == "Good":
                dates = ws.range(address=f"A{j-2}:A{j}")

                dates = [cleanup(l[0]) for l in dates]
       
                date = None
                wo_dt = None
                isblank = False
                if all([(len(l)<1 or l.isspace()) for l in dates]):
                    isblank= True
                
                for l in dates:  
                    if l!= "" and (not l.isspace()): print(l)
                    else:
                        continue

                    if bool(re.search('-|to|TO', l)) and any(char.isdigit() for char in l):
                        print("trying to add datetime")
                        date = re.findall(r'(?<!\/)\b\d{1,2}\/\d{1,2} \d{2}:\d{2}',l)
                        if len(date) > 0:
                            date = [std_date(n) for n in date]
                        
                        if len(date) == 0:
                            date = re.findall(r'\b\d{2}:\d{2} \d{1,2}/\d{1,2}',l)
                            if len(date) > 0:
                                date = [swap_date_time(m) for m in date]
                        
                        if len(date) == 0:
                            date = re.findall(r'\b\d{1,2}\/\d{1,2}\/\d{2,4} \d{2}:\d{2}', l)

                        print(date)
                        if any(elem in sheet for elem in ("25","24","23")):
                            try:
                                wo_dt = dt.datetime.strptime(date[0], "%d/%m/%y %H:%M")
                            except:
                                pass
                            try:
                                date = insert_after_first_slash(date[0], sheet[-2:])
                            except:
                                print(sheet)
                                print([m for m in dates])
                                print(date)
                                date = input("please type the date and time in the format DD/MM/YY HH:MM if it is not a date type '':")
                                if date== "":
                                    break
                                else:
                                    continue
                            try:
                                wo_dt = dt.datetime.strptime(date, "%d/%m/%y %H:%M")
                            except ValueError:
                                try:
                                    wo_dt = dt.datetime.strptime(date, "%H:%M %d/%m/%y")
                                except ValueError:
                                    print(sheet)
                                    print([m for m in dates])
                                    print(date)
                                    # wo_dt = dt.datetime.strptime(input("please type the date in dd/mm/yy HH:MM format:"), "%d/%m/%y %H:%M")
                                    while True:
                                        date = input("Please type the date and time if available in the following format DD/MM/YY HH:MM\n :")
                                        if date == "":
                                            break
                                        try:
                                            wo_dt = dt.datetime.strptime(date, "%d/%m/%y %H:%M")
                                            break
                                        except:
                                            print("please enter the date in the correct format or enter no value to move on")                            
                            break

                        elif "earlier" in sheet:

                            try:
                                wo_dt = dt.datetime.strptime(date[0], "%d/%m/%y %H:%M")
                            except:
                                print (f"The Program is struggling to interpret this data from sheet {sheet}" + "\n", dates)
                                while True:
                                    strings = " ".join(dates)
                                    working_data = re.findall(r"\d+\/\d+ \d{2}:\d{2}" ,strings)
                                    for indices , date_strings in enumerate(working_data):
                                        print( str(indices) + " [" + date_strings + "]" )
                                    date = input("Please type the date and time if available in the following format DD/MM/YY HH:MM\n :")
                                    if date == "":
                                        break
                                    if date in {"21","22","23"}:
                                        day_months, hours_mins = working_data[0].split(" ")
                                        day , month = day_months.split("/")
                                        if len(day) == 1:
                                            day = "0"+day
                                        if len(month)==1:
                                            month ="0"+month
                                        day_months = day+"/"+month
                                        new_date = day_months+"/"+ date +" " +hours_mins
                                        print(new_date)
                                        wo_dt = dt.datetime.strptime(new_date, "%d/%m/%y %H:%M")
                                        break

                                    try:
                                        wo_dt = dt.datetime.strptime(date, "%d/%m/%y %H:%M")
                                        break
                                    except:
                                        print("please enter the date in the correct format or enter no value to move on")                            
                            
                            break

                if date == None and isblank== False:
                    print (f"The Program is struggling to interpret this data from sheet {sheet}" + "\n", dates)
                    date = input("Please type the date and time if available in the following format DD/MM/YY HH:MM\n :")
                    if date != "":
                        wo_dt = dt.datetime.strptime(date, "%d/%m/%y %H:%M")

                data = ws.range(address=f"D{j-1}:S{j+10}")
                # Parse the data to generate a Work order
                if data[0][0]=="":
                    continue

                tonnage = data[2][9]
                if tonnage == "":
                    tonnage = data[2][5]
                if tonnage == 0:
                    continue
                wono = data[0][0]
                grade = data[1][0]
                raw_mat = {data[4:9][m][0]:data[4:9][m][5] for m in range(len(data[4:9]))}
                raw_mat_bom = {data[4:9][m][0]:data[4:9][m][12] for m in range(len(data[4:9]))}
                try:
                    # print ("removing addition "+str(data[8][10]))
                    raw_mat["RMSTDSAWDUST"] = raw_mat["RMSTDSAWDUST"] - data[8][10]
                except:
                    # print("couldn't remove sawdust correction")
                    pass

                hours = data[10][4]
                bom_hours = data[10][-4]
                
                # print("--------")
                
                try:
                    WO_list.append(MainPlantWorkOrder(tonnage,wono, grade, raw_mat, raw_mat_bom, hours, bom_hours, wo_dt))
                
                except:
                    print(data)
                    print("--------")
                    print(grade + " -- " + str(type(grade)))
                    print(str(tonnage) + " KG -- " + str(type(tonnage)))
                    print(raw_mat)
                    print(raw_mat_bom)
                    print(str(hours) + " hrs -- " + str(type(hours)))
                    print(str(bom_hours) + " hrs -- " + str(type(bom_hours)))
                    print("--------")
                    continue
                # WO_list.append(MainPlantWorkOrder(tonnage,wono, grade, raw_mat, raw_mat_bom, hours, bom_hours, wo_dt))
        report[sheet] = len(WO_list) - init_wo_length
        print(str(len(WO_list) - init_wo_length) + " WOs collected from " + sheet)

    print(str(len(WO_list)) + " work orders collected")
    print(report)

    new_db = xl.Database()


    new_db.add_ws(ws="WO_table")

    new_db.ws(ws="WO_table").update_index(row=1, col=1, val="Workorder number") 
    new_db.ws(ws="WO_table").update_index(row=1, col=2, val="Grade") 
    new_db.ws(ws="WO_table").update_index(row=1, col=3, val="Packed Weight") 
    new_db.ws(ws="WO_table").update_index(row=1, col=4, val="Hours per ton") 
    new_db.ws(ws="WO_table").update_index(row=1, col=5, val="Std Hours per ton")
    new_db.ws(ws="WO_table").update_index(row=1, col=6, val="datetime")
    for i ,j in enumerate(WO_list[0].ratios):
        new_db.ws(ws="WO_table").update_index(row=1, col=7+2*i, val=j + " ratio")
    for i , j in enumerate(WO_list[0].std_ratios):
        new_db.ws(ws="WO_table").update_index(row=1, col=8+2*i, val=j+ " std ratio")

    for row_id, data in enumerate(WO_list, start=1):
        new_db.ws(ws="WO_table").update_index(row=row_id+1, col=1, val=data.wono) 
        new_db.ws(ws="WO_table").update_index(row=row_id+1, col=2, val=data.grade) 
        new_db.ws(ws="WO_table").update_index(row=row_id+1, col=3, val=data.tonnage) 
        new_db.ws(ws="WO_table").update_index(row=row_id+1, col=4, val=data.route) 
        new_db.ws(ws="WO_table").update_index(row=row_id+1, col=5, val=data.std_route) 
        new_db.ws(ws="WO_table").update_index(row=row_id+1, col=6, val=data.datestr) 
        for i ,j in enumerate(data.ratios):
            new_db.ws(ws="WO_table").update_index(row=row_id+1, col=7+2*i, val=data.ratios[j])
        for i , j in enumerate(data.std_ratios):
            new_db.ws(ws="WO_table").update_index(row=row_id+1, col=8+2*i, val=data.std_ratios[j])
    
    xl.writexl(db=new_db, fn="WO_table_2506.xlsx")
    os.system("start EXCEL.EXE WO_table_2506")
        

if __name__=="__main__":
    main()



    
