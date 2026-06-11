import os
import sqlite3
import pandas as pd
import datetime as dt

script_dir = os.path.dirname(os.path.abspath(__file__))
db_file_path = os.path.join(script_dir, 'UF.db')

# Connect to or create the database
conn = sqlite3.connect(db_file_path)
# Create a cursor object to execute SQL commands
cursor = conn.cursor()
# Check if the 'books' table already exists, if not, create it
cursor.execute('''
               CREATE TABLE IF NOT EXISTS movements (
               id INTEGER PRIMARY KEY,
               Datetime DATETIME,
               Membrane_ID TEXT ,
               Action TEXT,
               Location TEXT,
               Position TEXT,
               Comment TEXT,
               Integrity TEXT,
               Permeability FLOAT,
               Tubes_sealed INTEGER,
               Report_ref TEXT,
               Pre_clean_notes TEXT,
               Post_clean_notes TEXT
               )
               ''')
cursor.execute('''
               CREATE TABLE IF NOT EXISTS membranes (
               purchase_id INTEGER PRIMARY KEY,
               Membrane_ID TEXT,
               DOM DATETIME
               )
               ''')
# Commit the changes and close the connection

conn.commit()
conn.close()

conn = sqlite3.connect(db_file_path)
# sqlite_tabel_data = pd.read_sql_query("SELECT * FROM movements", conn)


excel_file = pd.ExcelFile("M:\\PROCESS\\26. MBR\\MBR membrane tracking\\Membrane Tracker_2022.xlsx")
excel_sheet_name = 'Tracker'
excel_data = excel_file.parse(excel_sheet_name, header=1)

membrane_data = excel_file.parse("Look ups")
membrane_data = membrane_data.rename(columns={"Membrane IDs":"Membrane_ID", "Year of manufacture":"DOM"})
membrane_data = membrane_data[["Membrane_ID", "DOM"]]
print(membrane_data)
membrane_data.to_sql("membranes", conn, if_exists="append", index=False)

excel_data['Datetime'] = [pd.Timestamp.combine(row['Date'].date(), row['Time']) for _, row in excel_data.iterrows()]

excel_data = excel_data.rename(columns={"Membrane ID":"Membrane_ID", "UF bay":"Location", "UF position":"Position","Notes Additional":"Comment","Tubes sealed":"Tubes_sealed", "Notes from report pre clean":"Pre_clean_notes", "Notes from report post clean":"Post_clean_notes", "Report reference (N/A if none)":"Report_ref"})

sql_ready_table = excel_data[["Datetime", "Membrane_ID", "Action", "Location", "Position", "Comment", "Integrity", "Permeability","Tubes_sealed","Report_ref","Pre_clean_notes","Post_clean_notes"]]

sql_ready_table.to_sql("movements", conn, if_exists='append', index=False)

print(sql_ready_table)

conn.commit()
conn.close()
