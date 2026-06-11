import win32com.client
from datetime import datetime as dt
from datetime import timedelta as td
import pylightxl as xl
import xlsxwriter as xlwr
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl import utils
import string

class CustomError(Exception):
    def __init__(self, message):
        super().__init__(message)
        self.message = message

class Status:
    def __init__(self, asset, start, stop, duration, status):
        self.asset = asset
        self.start = dt.strptime(start,"%d-%b-%y %H:%M:%S")
        self.stop = dt.strptime(stop,"%d-%b-%y %H:%M:%S")
        self.delT = round(float(duration),4)
        self.status = status
    
    def __eq__(self,other):
        if not isinstance(other, Status):
            return NotImplemented
        return self.__dict__ == other.__dict__

    def __hash__(self):
        return hash(tuple(self.__dict__))
    
    def __add__(self,other):    
        
        if self.asset != other.asset:
            raise CustomError("These Status records are not for the same asset")
        
        if self.status != other.status:
            raise CustomError("These Status records are not of the same staus")
        
        if not(((self.start - other.stop).total_seconds() == 0.0) or ((self.stop - other.start).total_seconds() == 0.0)):
            difference_1 = self.start - other.stop
            difference_2 = self.stop - other.start
            # print(f"Neither {difference_1.total_seconds()} or {difference_2.total_seconds()} are 0.0") 
            raise CustomError("These Status Records are not adjacent and so cannot be combined")
         
        start = min(self.start, other.start)
        stop = max(self.stop, other.stop)
        delT = self.delT + other.delT
        return Status(self.asset, start.strftime("%d-%b-%y %H:%M:%S"), stop.strftime("%d-%b-%y %H:%M:%S"), delT, self.status)

    def __str__(self):
        return f"printing Status data {self.asset} {self.start.isoformat()} {self.stop.isoformat()} {self.delT} {self.status}"
    

def strip_msg_data(msg_string):
    list_of_lines = msg_string.splitlines()
    status_lines = []
    status_objects = []
    for i in list_of_lines:
        if i.endswith("Shutdown") or i.endswith("Makeload"):
            status_lines.append(i)
    
    for i in status_lines:
        status_array = i.rsplit(" ")
        status_array[:] = [x for x in status_array if x]
        if len(status_array) ==8:
            status_array[0] = status_array[0] + " " + status_array[1]
            status_array.pop(1)
        
        data = Status(status_array[0], status_array[1] + " " + status_array[2], status_array[3]+ " " + status_array[4], status_array[5], status_array[6])
        status_objects.append(data)

    return status_objects

# Function to recursively print folder content
def list_folders(folder, indent=0):
    print("\t  " * indent + folder.Name)
    for subfolder in folder.Folders:
         list_folders(subfolder, indent + 1)


def sum_statuses(list_of_statuses):
    sorted_list = sorted(list_of_statuses, key=lambda x: x.start, reverse=True)
    print(f"Summing {list_of_statuses[0].asset} statuses")
    count_of_sumation = 0
    new_list = []
    
    i = 1
    sumed_status = sorted_list[0]
    while i < len(sorted_list):
        try: 
            sumed_status += sorted_list[i]
            count_of_sumation+= 1
            i+=1

        except CustomError:
            new_list.append(sumed_status)
            sumed_status = sorted_list[i]
            i+=1
            continue
    new_list.append(sumed_status)
    
    # print(new_list[-1].start.strftime("%Y-%m-%d %H:%M:%S") + " " + new_list[-1].stop.strftime("%Y-%m-%d %H:%M:%S"))
    # print(sumed_status.start.strftime("%Y-%m-%d %H:%M:%S")+ " " + sumed_status.stop.strftime("%Y-%m-%d %H:%M:%S"))
    print(new_list[-1])
    print(sumed_status)

    print(f"Sorted list length {len(sorted_list)}; summed list length {len(new_list)}")
    print(f"We summed {count_of_sumation} statuses")
    return new_list

def Getletterfromindex(num):
    #produces a string from numbers so

    #1->a
    #2->b
    #26->z
    #27->aa
    #28->ab
    #52->az
    #53->ba
    #54->bb

    num2alphadict = dict(zip(range(1, 27), string.ascii_lowercase))
    outval = ""
    numloops = (num-1) //26
    
    if numloops > 0:
        outval = outval + Getletterfromindex(numloops)
        
    remainder = num % 26
    if remainder > 0:
        outval = outval + num2alphadict[remainder]
    else:
        outval = outval + "z"
    return outval

def create_sheet(statuses):
  
    # db = xl.Database()
    # workbook = xlwr.Workbook("plant_email.xlsx")
    # date_format = workbook.add_format({'num_format':"yyyy-mm-dd hh:mm:ss"})
    op_wb = Workbook()
    del op_wb['Sheet']

    assets = { 
        "Kiln 1":sum_statuses([x for x in statuses if x.asset== "Kiln 1"]),
        "Blend":sum_statuses([x for x in statuses if x.asset== "Blend"]),
        "Dryer 1":sum_statuses([x for x in statuses if x.asset== "Dryer 1"]),
        "MBR":sum_statuses([x for x in statuses if x.asset== "MBR"]),
        "Mill 1":sum_statuses([x for x in statuses if x.asset== "Mill 1"]),
        "Mill 2":sum_statuses([x for x in statuses if x.asset== "Mill 2"]),
        "Pelletisin":sum_statuses([x for x in statuses if x.asset== "Pelletisin"]),
        "Screening":sum_statuses([x for x in statuses if x.asset== "Screening"]),
        "Tank Farm":sum_statuses([x for x in statuses if x.asset== "Tank Farm"]),
        "Wash 1":sum_statuses([x for x in statuses if x.asset== "Wash 1"]),
        "Wood 1":sum_statuses([x for x in statuses if x.asset== "Wood 1"])
    }

    for i in assets:
        # db.add_ws(ws=i)
        # workbook.add_worksheet(i)
        op_wb.create_sheet(title=i)


    for i in assets:
        
        # for col_number , headings in enumerate(assets[i][0].__dict__):
        #     db.ws(ws=i).update_index(row=1,col=1+col_number, val=headings)
        #     workbook.get_worksheet_by_name(i).write(0,col_number,headings)
        
        # Headings for openpyxl workbook
        op_wb[i].append([x for x in assets[i][0].__dict__])

        for row_counter, j in enumerate(assets[i]):
            for col_index , attribute in enumerate(j.__dict__):
                if isinstance(j.__dict__[attribute], dt):
                    # db.ws(ws=i).update_index(row=2+row_counter,col=1+col_index, val=j.__dict__[attribute].strftime("%Y-%m-%d %H:%M:%S"))
                    # workbook.get_worksheet_by_name(i).write_datetime(row_counter+1,col_index,j.__dict__[attribute],date_format)
                    cell = op_wb[i].cell(row=row_counter+2,column=col_index+1,value=utils.datetime.to_excel((j.__dict__[attribute])))
                    cell.number_format = "YYYY-MM-DD hh:mm:ss"
                
                else:
                    # db.ws(ws=i).update_index(row=2+row_counter,col=1+col_index, val=j.__dict__[attribute])
                    # workbook.get_worksheet_by_name(i).write(row_counter+1,col_index,j.__dict__[attribute])
                    cell = op_wb[i].cell(row=row_counter+2,column=col_index+1,value=j.__dict__[attribute])
                    cell.number_format = 'General'
        
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab = Table(displayName=i.replace(" ","_"), ref = f"A1:{Getletterfromindex(len(assets[i][0].__dict__))}{len(assets[i])+1}")
        tab.tableStyleInfo = style
        op_wb[i].add_table(tab)
    
    # xl.writexl(db=db,fn="plant_email_output.xlsx")
    # workbook.close()
    op_wb.save("openpyxl_output.xlsx")


def Update_sheet(statues):

    assets = { 
        "Kiln 1":sum_statuses([x for x in statuses if x.asset== "Kiln 1"]),
        "Blend":sum_statuses([x for x in statuses if x.asset== "Blend"]),
        "Dryer 1":sum_statuses([x for x in statuses if x.asset== "Dryer 1"]),
        "MBR":sum_statuses([x for x in statuses if x.asset== "MBR"]),
        "Mill 1":sum_statuses([x for x in statuses if x.asset== "Mill 1"]),
        "Mill 2":sum_statuses([x for x in statuses if x.asset== "Mill 2"]),
        "Pelletisin":sum_statuses([x for x in statuses if x.asset== "Pelletisin"]),
        "Screening":sum_statuses([x for x in statuses if x.asset== "Screening"]),
        "Tank Farm":sum_statuses([x for x in statuses if x.asset== "Tank Farm"]),
        "Wash 1":sum_statuses([x for x in statuses if x.asset== "Wash 1"]),
        "Wood 1":sum_statuses([x for x in statuses if x.asset== "Wood 1"])
    }

    wb_obj = openpyxl.load_workbook("openpyxl_output.xlsx")




# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the folder (e.g., Inbox)
folder = outlook.Folders("andrew.carruthers@norit.com").Folders("Plant")

# Get the items in the folder
items = folder.Items

statuses = []
# List the subjects of the emails
for item in items:
    # print(item.body)
    statuses+= strip_msg_data(item.body)
# print(statuses)

create_sheet(statuses)





