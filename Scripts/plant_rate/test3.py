import pyodbc

connection_string = 'DSN=GLdata'
try:
    con = pyodbc.connect(connection_string)
    print("Connection successful!")
except pyodbc.Error as e:
    print(f"Connection failed: {e}")