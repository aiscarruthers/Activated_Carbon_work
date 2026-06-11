import os
import xlsxwriter as xlw
import pylightxl as xl
import json

def main():

    file = xl.readxl("Cmill.xlsx")
    sheet = file.ws("data")
    data_logger = {}

    def lims_to_tol(mean, UCL, units):
        if isinstance(mean, str):
            return "nonsense"
        tol = UCL - mean
        info = {
            "mean": round(mean,2),
            "tol": round(tol,2),
            "units": units
        }
        return info
    
    for row in sheet.rows:
        info = lims_to_tol(row[2],row[3],row[6])
        if row[0][-2:] == "PV":
            key = row[1][:4]
            key = key.replace(" ","")
            data_logger[key] = info
        elif row[0][-2:] == "OP":
            key = row[1][:4]
            key = key.replace(" ","")
            key = key + "_OP"
            data_logger[key] = info
        elif row[0][-3:] == "DCS":
            if row[7] == 'Invalid':
                data_logger["recipe"] = input("Recipe?: ")
            else:
                data_logger["recipe"] = str(int(round(float(row[7]))))
            
    
    data_logger["date"] = input("date(YYYY-MM-DD): ")

    print(data_logger)
    def compare_to_recipe(data):

        with open("Milling_recipes.json", "r") as recipe_file:
            recipe_settings = json.load(recipe_file)
        
        settings = recipe_settings[data["recipe"]]

        data["Grade"] = recipe_settings[data["recipe"]]["Grade"]

        for i in settings:
            if i in data and i != "Grade":
                print(data[i])
                data[i]["set point"] = settings[i]

        return data

    data_logger = compare_to_recipe(data_logger)

    workbook = xlw.Workbook("C_mill_entry.xlsx")
    report = workbook.add_worksheet("report")
    counter = 0
    for i in data_logger:
        intermediate = data_logger[i]    
        if isinstance(intermediate, dict):
            for j in intermediate:
                report.write(0,counter, i + " " + j)
                counter += 1
        else:
            report.write(0,counter, i)
            counter += 1
    
    counter = 0
    for i in data_logger:
        intermediate = data_logger[i]    
        if isinstance(intermediate, dict):
            for j in intermediate:
                report.write(1,counter, intermediate[j])
                counter += 1
        else:
            report.write(1,counter, intermediate)
            counter += 1

    workbook.close()

    os.system("start EXCEL.EXE C_mill_entry.xlsx")

if __name__=="__main__":
    main()
