import sys
import os
import warnings
import xlsxwriter as xlw
import pylightxl as xl
import glob
import datetime
import re
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import json
import cProfile
warnings.filterwarnings(action='ignore', category=UserWarning)

# Dictionary of the grades and the recipes they correspond to not sure how accurate this is
recipe_file = os.path.dirname(os.path.realpath(__file__)) + "/recipebygrade.json"
with open(recipe_file) as f:
    recipie_dict = json.load(f)

class Pallet():
    """Pallet class, this stores all the information you might need for your 
    pallets, except perhaps where they are in real life.
    
    Attributes:
        number (int): Ideally from 1 to 40 but values above are not uncommon 
        especially when manufacturing Dryer Product or other C-carbons.
        pln (int):  Pallet Lot Number is a unique seven digit integer.
        dom (datetime obj): Date of manufacture, Datetime obj of when the 
        pallet was packed.
        grade (string): This string indicates the Grade of the pallet, common
        grades include CNR115, CNR115lb, CN120, CA1 etc.
        quality (string): Quality of the pallet, is this alright to ship, if 
        yes string is "FPQ", else "OQ".
        weight (float): the weight in tonnes of the pallet
        shift (string): Day shift "D" or night shift "N"
        team (string): The team of operators working while the pallet was 
        packed
        comments (string): Comments made on the pallet

    """

    def __init__(self, num, pln, dom, grd, wt, shft, team):
        """Pallet constructor to initialise this object
        
        Args:
            num (int): ideally from 1 to 40 but values above are not uncommon
            pln (int): Pallet Lot Number is a unique seven digit integer.
            dom (string): Date of manufacture, this is a date time object of 
            when the pallet was packed. Care should be taken as to the date 
            format being read into the program. All dates out are given in ISO 
            standard.
            grd (string): This string indicates the Grade of the pallet, common
            grades include CNR115, CNR115lb, CN120, CA1PA etc.
            wt (float): the weight in tonnes of the pallet
            shft (string): Day shift "D" or night shift "N"
            t (string): The team of operators working while the pallet was 
            packed
                    
        """
        self.number = num
        self.pln = pln
        self.dom = dom
        self.grade = grd
        self.quality = "FPQ"
        self.weight = wt
        self.shift = shft
        self.team = team
        self.recipe = ""
        self.comments = ""

        try:
            self.recipe = recipie_dict[grd]
        except KeyError:
            recip = input("What recipe is this: '{}'?:".format(grd))
            self.recipe = recip
            recipie_dict[grd] = recip

    def set_pln(self, pln):
        "Set the Pallet Lot Number."
        self.pln = pln

    def get_pln(self):
        "Get the Pallet Lot Number."
        return self.pln

    def set_number(self, num):
        "Set the Pallet Number."
        self.number = num
    
    def get_number(self):
        "Get the Pallet Number."
        return self.number
    
    def set_dom(self, dom):
        "Set the Date Of Manufacture."
        self.dom = dom
    
    def get_dom(self):
        "Get the Date Of Manufacture."
        return self.dom
    
    def set_grade(self, grd):
        "Set the Grade of the Pallet."
        self.grade = grd
    
    def get_grade(self):
        "Set the Grade of the Pallet."
        return self.grade
    
    def set_quality(self, ql):
        "Set the Quality of the Pallet."
        self.quality = ql
    
    def get_quality(self):
        "Get the Quality of the Pallet."
        return self.quality

    def set_weight(self, wt):
        "Set the Weight of the Pallet."
        self.weight = wt 

    def get_weight(self):
        "Get the Weight of the Pallet."
        return self.weight 

    def set_shift(self, shft):
        "Set the type of Shift while the pallet was made, day 'D' or night 'N'."
        self.shift = shft

    def get_shift(self):
        "Get the type of Shift while the pallet was made, day 'D' or night 'N'."
        return self.shift

    def set_team(self, team):
        "Set the Team working while the Pallet was made."
        self.team = team

    def get_team(self):
        "Get the Team working while the Pallet was made."
        return self.team

    def set_comments(self, coms):
        "Set some Comments on the Pallets"
        self.comments = coms

    def get_comments(self):
        "Get some Comments on the Pallets"
        return self.comments

def day_night(path_to_shift_log):
    """Takes an xl shiftlog path, returns a US date and 2 worksheet objects"""
    # figuring out dates is a nightmare, Below is a database of how months 
    # should be interpreted from the file naming scheme. Woe betide those 
    # filthy americans and their topsyturvy date system.

    Dates = {".01." : "01", ".02.": "02", ".03.": "03", ".04.": "04",
    ".05.": "05", ".06.": "06", ".07.": "07", ".08.": "08", ".09.": "09",
    ".10.": "10", ".11.": "11", ".12.": "12", ".1.": "01", ".2.": "02", 
    ".3.": "03", ".4.": "04", ".5.": "05", ".6.": "06", ".7.": "07", 
    ".8.": "08", ".9.": "09"}
    
    # print(path_to_shift_log)
    file = xl.readxl(path_to_shift_log )
    
    # Here we are using the name of the shift log to identify the date. 
    # Note: if the shiftlogs names include additional conent within them this
    # should be no issue, however be carful of double counts if there are 
    # multiple logs of the same date. 
    # Additionally this will break if the day is given with 
    # only one number or it is given in an american date format.  
    
    def get_closest_matches(user_input, choices, threshold=70):
        """Get close matches to the user input from a list of choices."""
        matches = process.extract(user_input, choices, limit=None)
        return [match[0] for match in matches if match[1] >= threshold]

	# This function was required because certain shifts have figured out how to
    # corrupt the shiftlog excel files. Occasionally a person in operations that
    # is somewhat tech savy recovers them and notes the name of the shift log 
    # under a slightly different nameing convention.
    def pick_sheet(initial_sheet_name):
        """
        This function is used to accomodate situations where the sheetnames 
        have been changed. This allows the user to select a new sheet.

        Parameters:
        initial_sheet_name (string): This is the initial sheet name that is 
        being used to select a sheet and is used to find similarly named sheets

        Returns:
        Selected Worksheet (pylightxl worksheet object) 

        """
        sheet_names = file.ws_names
        close_matches = get_closest_matches(initial_sheet_name, sheet_names)

        if not close_matches:
            print("No close matches found.")
            return

        if len(close_matches) == 1:
            selected_sheet = close_matches[0]

        elif fuzz.ratio(close_matches[0], initial_sheet_name) == 100:
            selected_sheet = close_matches[0]
        
        else:
            print("\nLooks like some one has changed the sheet names:")
            for idx, sheet_name in enumerate(sheet_names):
                print(f"{idx + 1}. {sheet_name}")
            
            print("\nClose matches to {}:".format(initial_sheet_name))
            for idx, match in enumerate(close_matches):
                print(f"{idx + 1}. {match}")
            
            choice = int(input("\nEnter the number corresponding to the sheet you want to use: "))
            if 1 <= choice <= len(close_matches):
                selected_sheet = close_matches[choice - 1]
            else:
                print("Invalid choice.")
                return

        return file.ws(selected_sheet)

    day = pick_sheet("DAY_SHIFT")
    night = pick_sheet("NIGHT_SHIFT")
    reprocessed = pick_sheet("WEIGHT,DISP,REPROC") # changed from reprocessing on the 21st May 2024
    
    # The code block below was developed during the IT hand over to Norit. 
    # Operators were for a period unable to access shiftlogs and would change 
    # the sheet names after the fact to things like, 'Dayshift recovered' this
    # would throw the program into an error. I tried to develop some 
    # functionality to get around this problem using regular expression 
    # searching of the sheet names. But this made the program unnessearily slow 
    # the pick_sheet function addresses this problem very well

    for i in Dates:
        if i in path_to_shift_log:
            leng = len(i)
            x = path_to_shift_log.index(i)
            year_end =  re.search(r'[\W\D]', path_to_shift_log[x+leng:])
            DOM = path_to_shift_log[ x-2:x] + "/" + Dates[i] + "/" + path_to_shift_log[x+leng:x+leng+year_end.start()]
            continue
    
    if len(DOM) == 10:
        DOM = datetime.datetime.strptime(DOM, "%d/%m/%Y" )
    elif len(DOM) == 8:
        DOM = datetime.datetime.strptime(DOM, "%d/%m/%y" )
    # print(DOM)
    return day, night, reprocessed, DOM

def pull_sheet(sheet, shift , date):
    """ Takes a shift worksheet and generates a list of pallet objects

    Parameters:
    sheet (worksheet obj): This is a woksheet object of a day or night shift.
    shift (string): Was the shift day, "D", or night, "N"?
    date (string): A United States date string.

    Returns:
    lots (list): This is a list of unique pallet objects
    
    """
    # Gen empty list to be filled and pulls the team letter
    lots = []
    Team = sheet.index(row=2, col = 6)

    day_night_flag = 0
    if shift == "D":
        day_night_flag = 1
    elif shift =="N":
        day_night_flag = 0
    for i in range(4, 11):
        PLN = sheet.index(row = 15 + day_night_flag, col = i)
        Grade = sheet.index(row = 16 + day_night_flag, col = i)
        start = sheet.index(row = 17 + day_night_flag, col = i)
        stop = sheet.index(row= 18 + day_night_flag, col = i)
        Weight = sheet.index(row = 20 + day_night_flag, col = i)
        
        if Weight != "":
            if type(Weight) is int:
                try:
                    Weight = Weight / 1000
                except TypeError:
                    print(f"we cant do maths to the weight of PLN {str(PLN)} from the {shift} shift of {date.date()}: {Weight}")
            else:
                try:
                    Weight = int(Weight.strip())/1000
                except:
                    print(f"we cant do maths to the weight of PLN {str(PLN)} from the {shift} shift of {date.date()}: {Weight}")

        # if CP2 have written a start value the same as a stop value this tends
        # to mean there was only one pallet finished
        if start == stop and start != "":
            lots.append(Pallet( start, PLN, date, Grade, Weight, shift, Team))

        elif start != "" and stop != "":
            try: 
                start = int(start)
                stop = int(stop)
            except ValueError as e:
                lots.append(Pallet( start, PLN, date, Grade, Weight, shift, Team))
                print ("PLN {} packed on {} shift on the {} may not be on a pallet".format(str(PLN), shift, date.date()))
                continue
            # if the start is after the end idk what this means so it throws a 
            # warning to the user
            if start > stop:
                print("Look out PLN {} ended before it began on the {} shift of {}".format(str(PLN), shift, date.date()))
                continue
            
            # If there is a start value and a stop value then the code runs
            else:
                for j in range(start, stop + 1):
                    lots.append(Pallet(j, PLN, date, Grade, Weight, shift, Team))
            
        # for every other condition it skips to the next column. Like if a 
        # Start value is given but there is no stop value this tends to mean 
        # the pallet was started but has not been finished. usually this is 
        # caught the next shift

        else:
            continue

    return lots

def ps_wrapper(lst, sheet, shift, date):
    """This is a wrapper for the pull_shift function which would allow it to be 
    used over a series of sheets or just one to then append it's results to an external list
    """
    for pallets in pull_sheet(sheet, shift, date):
        lst.append(pallets)

def pull_tankers(sheet, DOM):
    """Takes the sheet object and pulls the tanker movements into a list"""
    tankers = []
    shift = sheet.address(address="F2")
    cells = sheet.col(2)
    start_index = [i for i, strings in enumerate(cells) if "TANKER MOVEMENTS" in str(strings)][-1]
    stop_index = [i for i, strings in enumerate(cells) if "DOWNTIME / PRODUCTION RATE" in str(strings)][-1]
	
    for i in range(start_index + 3, stop_index):
        tanker_mvmt = sheet.row(i)[1:]
        if all([i == "" for i in tanker_mvmt]) == True:
            continue
        # print(len(tanker_mvmt))
        tanker_mvmt.pop(2)
        tanker_mvmt = tanker_mvmt[:10]
        # print(len(tanker_mvmt))

        tanker_mvmt.insert(0, DOM)
        tanker_mvmt.append(shift)

        tankers.append(tanker_mvmt)

    return tankers

def pT_wrapper(lst, sheet, DOM):
    """This is a wrapper for the pull_tankers function which would allow it to be 
    used over a series of sheets or just one to then append it's results to an external list
    """
    # print("pm_wrapper initiated")
    for tanker_mvmnts in pull_tankers(sheet,DOM):
        lst.append(tanker_mvmnts)


def write_to_Quality_xl(list_of_pallets, log_of_comments, Maint_WOs, Tanker_Mvmt,
                        file=os.path.dirname(os.path.realpath(__file__)) + "/shift_pull.xlsx"):
    """Takes a list of pallet objs, writes to an xlsx file given a path."""
    
    # Generates a XlsxWriter workbook object and uses this to write an xl file
    # idk how this works fully, but it does.

    workbook = xlw.Workbook(file)
    # adding cell formating to sheet
    cell_format = workbook.add_format({'bold':True})
    cell_format.set_bg_color('#92D050')
    cell_format.set_text_wrap()
    cell_format.set_align("center")
    cell_format.set_align("vcenter")
    prod = workbook.add_worksheet("Production")
    
    # creating the date fromat
    date_format = workbook.add_format({'num_format':'yyyy-mm-dd'})
    
    # headings of the production sheet
    prod.write(0, 0, "Date", cell_format)
    prod.write(0, 1, "Today", cell_format )
    prod.write(0 ,2, "Age", cell_format)
    prod.write(0, 3 , "Campaign", cell_format)
    prod.write( 0,  4, "Day/Night", cell_format)
    prod.write( 0,  5, "Team", cell_format)
    prod.write( 0, 6, "Recipe", cell_format)
    prod.write( 0, 7, "Grade", cell_format)
    prod.write( 0,  8, "PLN", cell_format)
    prod.write( 0,  9, "Weight", cell_format)
    prod.write( 0,  10,  "Number", cell_format)
    prod.write( 0,  11,  "Quality", cell_format)
    prod.write(0, 12, "502 503 entry Window", cell_format)
    prod.write(0, 13, "502 503 exit Window", cell_format)
    prod.write(0, 13, "Detractor", cell_format)
    prod.write(0, 14, "Failure Rootcause Category", cell_format)
    prod.write(0, 15, "Root Cause", cell_format)
    prod.write(0, 16, "Contributing factors", cell_format)
    prod.write( 0,  17,  "Comments", cell_format)

    # in this for loop we print out all the data from the pallets we generated,
    # skiping the first row (hence +1), into the xl file as we go down the rows.  
    for i , j in enumerate(list_of_pallets):
        i = i + 1
        prod.write_datetime(i, 0, j.dom.date(), date_format)
        prod.write_formula(i, 1, "=TODAY()", date_format)
        prod.write_formula(i,2, "=B{0}-A{0}".format(str(i+1)))
        prod.write( i,  4, j.shift)
        prod.write( i,  5, j.team)
        prod.write( i, 6, j.grade)
        prod.write( i, 7, j.recipe)
        prod.write( i,  8, j.pln)
        prod.write( i,  9, j.weight)
        prod.write( i,  10,  j.number)
        prod.write( i,  11,  j.quality)
    # adds filter so you don't have to
    prod.autofilter(0,0,len(list_of_pallets)+1, 10)

    com_format = workbook.add_format()
    com_format.set_text_wrap()
    
    log = workbook.add_worksheet("Log_comments")
    log.set_column(2,2, 100, com_format)
    log.write(0,0, "Date", cell_format)
    log.write(0,1,"Time", cell_format)
    log.write(0,2,"Comms", cell_format)
    
    for i, j in enumerate(log_of_comments):
        i = i + 1
        log.write_datetime(i, 0, j[0].date(), date_format)
        log.write(i, 1, j[1])
        log.write(i, 2, j[2])

    log.autofilter(0,0,len(log_of_comments)+1, 2)

    maint_log = workbook.add_worksheet("Reactive Maintenance")
    maint_log.write(0,0, "Date",cell_format)
    maint_log.write(0,1, "Time",cell_format)
    maint_log.write(0,2, "Duration",cell_format)
    maint_log.write(0,3, "Area",cell_format)
    maint_log.write(0,4, "Asset",cell_format)
    maint_log.write(0,5, "Description",cell_format)
    maint_log.write(0,6, "WO",cell_format)
    maint_log.write(0,7, "Contractor",cell_format)
    maint_log.write(0,8, "Shift",cell_format)

    for index , Wo in enumerate(Maint_WOs):
        i=index+1
        for k, entry in enumerate(Wo):
            if k ==0:
                maint_log.write(i,k,entry.date(),date_format)
            else:
                maint_log.write(i,k,entry)
    maint_log.autofilter(0,0,len(Maint_WOs)+1,8)

    tanker_log = workbook.add_worksheet("Tanker Movements")
    tanker_log.write(0,0, "Date",cell_format)
    tanker_log.write(0,1, "Haulier",cell_format)
    tanker_log.write(0,2, "Reg / TANK NO",cell_format)
    tanker_log.write(0,3, "Description",cell_format)
    tanker_log.write(0,4, "Time in",cell_format)
    tanker_log.write(0,5, "Time out",cell_format)
    tanker_log.write(0,6, "Weight in kg",cell_format)
    tanker_log.write(0,7, "Weight out kg",cell_format)
    tanker_log.write(0,8, "Net kg",cell_format)
    tanker_log.write(0,9, "Lot No",cell_format)
    tanker_log.write(0,10, "Comment",cell_format)

    for index , tanker in enumerate(Tanker_Mvmt):
        i=index+1
        for k, entry in enumerate(tanker):
            if k ==0:
                tanker_log.write(i,k,entry.date(),date_format)
            else:
                tanker_log.write(i,k,entry)
    tanker_log.autofilter(0,0,len(Tanker_Mvmt)+1,10)

    workbook.close()
    os.system(f"start EXCEL.EXE {file}")

def pull_prod(path_to_shift_log):
    """Takes the file path of a shift log, returns a list of pallet objects."""
    # Here we're pulling together all the previous functions for the script
    day, night, reprocessed, DOM = day_night(path_to_shift_log)
    day_lots = pull_sheet(day, "D", DOM)
    night_lots = pull_sheet(night, "N", DOM)
    lots = day_lots + night_lots
    return lots

def pull_comms(sheet, DOM):
    """Takes the sheet object and pulls the comunication entries into a list"""
    total_com = []
    sign_off_array = sheet.col(2)
    hand_over_array = sheet.col(3)
    sign_off_index = [i for i, strings in enumerate(sign_off_array) if "SM Sign Off" in str(strings)][-1]
    hand_over_index = [i for i, strings in enumerate(hand_over_array) if "MANAGER HANDOVER" in str(strings)][-1]
	
    for i in range(hand_over_index + 1, sign_off_index + 2):
        com = []
        com.append(DOM)
        com.append(sheet.index(row=i, col = 2))
        com.append(sheet.index(row=i, col = 3))
        total_com.append(com)
    
    return total_com

def pc_wrapper(lst, sheet, DOM):
    """This is a wrapper for the pull_coms function which would allow it to be 
    used over a series of sheets or just one to then append it's results to an external list
    """
    # print("pulling coms from sheet")
    for notes in pull_comms(sheet,DOM):
        lst.append(notes)

def pull_logs(path_to_shift_log):
    """ Takes the file path of a shift log and returns the comments of the day
    and night shift with a date"""
    day, night, reprocessed, DOM = day_night(path_to_shift_log)
    day_coms = pull_comms(day, DOM)
    night_coms = pull_comms(night, DOM)
    coms = day_coms + night_coms
    return coms

def pull_maintenance(sheet,DOM):
    """ Takes the sheet object of the shiftlog and pulls the entries in the 
    Maintenance log, filtering out empty rows of data and adding the Shift identifier
    """
    # print("pulling maintenance data from sheet")
    total_maintenance = []
    shift = sheet.address(address="F2")

    maintenance_start_index = [i for i, strings in enumerate(sheet.col(2)) if "MAINTENANCE" in str(strings)][-1]
    maintenance_end_index = [i for i, strings in enumerate(sheet.col(2)) if "CP2 HANDOVER" in str(strings)][-1]
    
    # print(shift, maintenance_start_index, maintenance_end_index)

    for i in range(maintenance_start_index+3,maintenance_end_index+1):
        maintenance_data = sheet.row(i)[1:]
        maintenance_data.pop(8)
        maintenance_data.pop(7)
        maintenance_data.pop(5)
        maintenance_data.pop(2)
        while len(maintenance_data)>7 and maintenance_data[-1] == "":
            maintenance_data.pop()    

        if all(elements == "" for elements in maintenance_data):
            continue
        
        maintenance_data.insert(0,DOM) 
        maintenance_data.append(shift)
        # print(maintenance_data)
        total_maintenance.append(maintenance_data)
    return total_maintenance

def pm_wrapper(lst, sheet, DOM):
    """This is a wrapper for the pull_maintenance function which would allow it to be 
    used over a series of sheets or just one to then append it's results to an external list
    """
    # print("pm_wrapper initiated")
    for maint_WOs in pull_maintenance(sheet,DOM):
        lst.append(maint_WOs)
 
def pull_maint_Wos(path_to_shift_log):
    """ Takes the file path of a shift log and returns the reactive WOs for that
    day"""
    day, night, reprocessed, DOM = day_night(path_to_shift_log)
    day_WOs = pull_maintenance(day, DOM)
    night_WOs = pull_maintenance(night, DOM)
    WOs = day_WOs + night_WOs
    return WOs


def pull_frm_args():
    """Takes this script's args, writes an xlsx file containing pallet info."""
    # This is the final function that will run when the script is called
    # it pulls in arguments from the command line using the sys library and 
    # pipes them into the pull_log function and then appends each pallet
    # into into the empty list before using that list in the final 
    # write_to_Quality_xl function which creates the xl file. 
    
    paths = sys.argv
    lots = []
    coms = []
    maint_WOs = []
    tankers = []
    for l in paths:
        try:
            if os.path.isfile(l) and l[-4:] in ("xlsx", "xlsm") and "~$" not in l:
                day , night, reprocessed, DOM = day_night(l)
                ps_wrapper(lots, day, "D",DOM)
                ps_wrapper(lots, night, "N",DOM)
                pT_wrapper(tankers, day, DOM)
                pT_wrapper(tankers, night, DOM)
                # psr_wrapper(repr, day, DOM)
                # psr_wrapper(repr, night, DOM)
                # print("running pc and pm wrappers")
                pc_wrapper(coms, day, DOM)
                pc_wrapper(coms, night, DOM)
                pm_wrapper(maint_WOs ,day, DOM)
                pm_wrapper(maint_WOs ,night, DOM)

                # repr.append(reprocessing_sheet_pull(reprocessed))
        
            elif os.path.isdir(l):
                for m in glob.glob("./"+ l + "/*"):
                    if m[-4:] in ("xlsx", "xlsm") and "~$" not in m: 
                        print(m)
                        day , night, reprocessed, DOM = day_night(m)
                        ps_wrapper(lots, day, "D",DOM)
                        ps_wrapper(lots, night, "N",DOM)
                        pT_wrapper(tankers, day, DOM)
                        pT_wrapper(tankers, night, DOM)
                        pc_wrapper(coms, day, DOM)
                        pc_wrapper(coms, night, DOM)
                        pm_wrapper(maint_WOs ,day, DOM)
                        pm_wrapper(maint_WOs ,night, DOM)
                        # for o in pull_prod(m):
                        #     lots.append(o)
                        # for o in pull_logs(m):
                        #     coms.append(o)
                        # for o in pull_maint_Wos(m):
                        #     maint_WOs.append(o)
        except :
            raise

    write_to_Quality_xl(lots, coms, maint_WOs, tankers)
# Now we run the final function.
# cProfile.run('pull_frm_args()', sort='cumtime')
pull_frm_args()

with open(recipe_file, "w") as f:
    f.write(json.dumps(recipie_dict, indent=2))
