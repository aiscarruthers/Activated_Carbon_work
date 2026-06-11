import datetime as dt
import json
import os
import sys
import glob
import pylightxl as xl
import re


class WorkOrder:
    def __init__(self, Order_number, grade, inputs, outputs, start, end):

        self.WO_no = Order_number
        self.grade = grade
        self.inputs = inputs
        self.output = outputs
        self.start = start
        self.end = end
        self.input_check()
    
    def input_check(self):
        for i in self.inputs:
            if type(self.inputs[i]) is str:
                print(f"{i} has been written as {self.inputs[i]} on the {self.start.isoformat()} what should this be")
                self.inputs[i] = float(input("Type a value: "))
    
    def __str__(self):
        strings = ""
        for i in self.inputs:
            if type(i) is str:
                var = self.inputs[i]
            else:
                var = round(self.inputs[i],0) 
            
            strings += f"\n {i}: {var}"        
        return f"PLN {self.WO_no} {self.start.isoformat()} {self.grade} {round(self.output,0)} kgs {strings}"

    def WO_comp(self, other):
        return isinstance(other, WorkOrder) and self.WO_no == other.WO_no
    
    def true_WO_comp(self,other):
        return isinstance(other, WorkOrder) and self.WO_no == other.WO_no and self.start == other.start and self.end == other.end

    def ne_WO_comp(self,other):
        return not self.WO_comp(other)

    def merge(self, otherWO):
        # print(self)
        # print(otherWO)
        if self.WO_no == otherWO.WO_no and not self.true_WO_comp(otherWO):
            for i in self.inputs:
                try:
                    self.inputs[i]+= otherWO.inputs[i]
                except KeyError:
                    print(f"problem merging {self.WO_no} from {self.start} and \n {otherWO.WO_no} from {otherWO.end}")
                    print(f"{i} of {self.inputs[i]}")
                    print("Printing input keys")
                    for j in self.inputs:
                        print(j)
                    print("Printing other inputs")
                    for j in otherWO.inputs:
                        print(j)
                    new_key = input("which key should be used instead if any: ")
                    other_key = input("what should the other key be: ")
                    if new_key == "0":
                        continue
                    else:
                        self.inputs[new_key] += otherWO.inputs[other_key]
   
            self.output += otherWO.output

            if self.start > otherWO.start:
                self.start = otherWO.start
            
            elif self.start < otherWO.start:
                self.end = otherWO.end
        else:
            return
    
    def to_dict(self):
        dictionary = {}
        dictionary["WO_no"] = self.WO_no
        dictionary["grade"] = self.grade
        dictionary["inputs"] = self.inputs
        dictionary["output"] = self.output
        dictionary["start"] = self.start.isoformat()
        dictionary["end"] = self.end.isoformat()
        return dictionary
    
    def write_to_json(self):
        script_directory = os.path.dirname(os.path.abspath(__file__))
        WO_filepath = os.path.join(script_directory, 'WO_data.json')
        try:    
            with open(WO_filepath, 'r') as json_file:
                data = json.load(json_file)
        except:
            data = {}

        if self.WO_no not in data:
            data[self.WO_no] = self.to_dict()
        
        elif self.WO_no in data:
            previous = WorkOrder.from_dict(data[self.WO_no])
            self.merge(previous)
            data[self.WO_no] = self.to_dict()
        
        with open(WO_filepath, "w") as json_file:
            json.dump(data, json_file, indent=4)            

    @classmethod
    def from_dict(cls, data):
        WO_no = data["WO_no"]
        grade = data["grade"] 
        inputs = data["inputs"] 
        output = data["output"] 
        start = dt.datetime.fromisoformat(data["start"])
        end = dt.datetime.fromisoformat(data["end"])
        return cls(WO_no, grade, inputs, output, start, end)


def gen_WOs(excel_path):
    print(excel_path)
    WOs = []

    file = xl.readxl(excel_path)
    sheet = file.ws("Summary")
    Date = dt.datetime.strptime(sheet.address(address="L4"), "%Y/%m/%d %H:%M:%S")

    for i in range(4,18):
        if sheet.index(row = 17, col = i) != "":
            WO_dict = {}
            inputs = {}
            for j in range(17,32):
                if j== 17:
                    WO_dict["WO_no"] = sheet.index(row=j,col=i)
                elif j == 18:
                    WO_dict["grade"] = sheet.index(row=j,col=i)
                elif j ==19:
                    WO_dict["output"] = sheet.index(row=j,col=i)
                elif sheet.index(row=j,col=i)!= "":
                    inputs[sheet.index(row=j, col=2)] = sheet.index(row=j,col=i)
            
            if type(sheet.index(row=16,col=i)) is not str:
                start_time = str(round(sheet.index(row=16, col=i),2))
            else:
                start_time = sheet.index(row=16, col=i)
            
            if start_time!= "":
                
                try:
                    time = dt.datetime.strptime(start_time, "%H:%M:%S").time()

                except:
                    print(start_time)
                    try:
                        time = dt.datetime.strptime(start_time, "%H.%M").time()
                    except ValueError:
                        print(sheet.index(row=16, col=i))
                        time = Date.time()
                    
                if  any([j == start_time for j in ["START" , "start" ,  "strt"]]):
                    WO_dict["start"] = Date.isoformat()
                
                elif any([j in start_time for j in ["START", "start", "strt"]]) and ":" in start_time:
                    
                    pattern = r'\d\d:\d\d|[^\n:]{2}:\d\d'
                    matches = re.findall(pattern, start_time)
                    
                    print(matches)
                    time = dt.datetime.strptime(matches[0], "%H:%M").time()
                    
                    if (time.hour + time.minute / 60) > dt.time(hour=7).hour:
                        WO_dict["start"] = dt.datetime.combine(Date.date(), time).isoformat()
                    elif (time.hour + time.minute / 60) < dt.time(hour=7).hour:
                        WO_dict["start"] = dt.datetime.combine((Date+dt.timedelta(days=1)).date(), time).isoformat()

                elif any([j in start_time for j in ["START", "start", "strt"]]) and "." in start_time:
                    
                    pattern = r'\d\d\.\d\d'
                    matches = re.findall(pattern, sheet.index(row=16,col=i))
                    
                    print(matches)
                    time = dt.datetime.strptime(matches[0], "%H.%M").time()
                    
                    if (time.hour + time.minute / 60) > dt.time(hour=7).hour:
                        WO_dict["start"] = dt.datetime.combine(Date.date(), time).isoformat()
                    elif (time.hour + time.minute / 60) < dt.time(hour=7).hour:
                        WO_dict["start"] = dt.datetime.combine((Date+dt.timedelta(days=1)).date(), time).isoformat()

                elif (time.hour + time.minute / 60) > dt.time(hour=7).hour:
                    WO_dict["start"] = dt.datetime.combine(Date.date(), time).isoformat()
                elif (time.hour + time.minute / 60) < dt.time(hour=7).hour:
                    WO_dict["start"] = dt.datetime.combine((Date+dt.timedelta(days=1)).date(), time).isoformat()
                else:
                    WO_dict["start"] = Date.isoformat()
            else:    
                WO_dict["start"] = Date.isoformat()

            try:
                WO_dict["end"] = (dt.datetime.fromisoformat(WO_dict["start"]) + dt.timedelta(hours=float(inputs["Hours"]))).isoformat()
            except:
                print(sheet.index(row=16,col=i))
                print(WO_dict["WO_no"])
                raise
            
            # print(WO_dict["end"])
            WO_dict["inputs"] = inputs
            WOs.append(WO_dict)             
    return WOs

def main_function():
    WOs = []
    objcts = []
    arguments = sys.argv
    for path in arguments:
        try:
            if os.path.isfile(path) and path[-4:] == "xlsx" and "~$" not in path:
                for i in gen_WOs(path):
                    WOs.append(i)
            
            elif os.path.isdir(path):
                for i in glob.glob("./"+path+"/*"):
                    if i[-4:]=="xlsx" and "~$" not in i:
                        for j in gen_WOs(i):
                            WOs.append(j)       
        except:
            raise
    for i in WOs:
        objcts.append(WorkOrder.from_dict(i))
    
    print(f"Length of collected WOs {len(objcts)}")
    merged_dict = {}
    final_list = []
    for instance in objcts:
        if instance not in merged_dict:
            
            duplicates = [other for other in objcts if instance.WO_comp(other)]

            for i, duplicate in enumerate(duplicates):
                if i != 0:
                    instance.merge(duplicate)

            for duplicate in duplicates:
                merged_dict[duplicate]= instance

            final_list.append(instance)
    
    for i in objcts:
        if any([j.WO_comp(i) for j in final_list]) == False:
            final_list.append(i)

    print(f"Length of Unique WOs {len(final_list)}")
    for i in final_list:
        i.write_to_json()
    

if __name__ == "__main__":
    main_function()
