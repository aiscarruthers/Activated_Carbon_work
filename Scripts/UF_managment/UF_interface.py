import tkinter as tk
# from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import sqlite3
import os
import datetime as dt
from ttkthemes import ThemedTk

# Get the directory of the script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the path to the database file
db_file_path = os.path.join(script_dir, 'UF.db')

# Function to display the current location of Membranes
def display_ufs():
    conn = sqlite3.connect(db_file_path)
    cursor = conn.cursor()
    # cursor.execute("SELECT * from movements;")
    cursor.execute('''SELECT m.Membrane_ID, m.Location, m.Position , m.Datetime AS latest_location,
                   p.DOM 
                   FROM ( SELECT 
                   Membrane_ID,
                   MAX(Datetime) AS latest_movement_date
                   FROM
                   movements
                   GROUP BY
                        Membrane_ID) AS latest_movement
                   JOIN
                        movements AS m ON latest_movement.Membrane_ID = m.Membrane_ID AND latest_movement.latest_movement_date= m.Datetime
                   JOIN
                        membranes as p on m.Membrane_ID = p.Membrane_ID;
                   ''')
    
    books = cursor.fetchall()
    conn.close()

    # Clear the listbox
    uf_listbox.delete()
    # print(books)
    # Insert books into the listbox
    for i, book in enumerate(books):
        if book[-1]!= "xx":
            age = str(round((dt.datetime.now() - dt.datetime.strptime(book[-1], "%Y-%m-%d %H:%M:%S")).days / 365, 2))
        else:
            age = "Unkown"

        time_in_location = str(round((dt.datetime.now() - dt.datetime.strptime(book[3],"%Y-%m-%d %H:%M:%S")).days, 2))
        uf_listbox.insert('', tk.END, values=( book[1], book[0], book[2], time_in_location, age))

# Function to add a movement to the database
def add_mvmt():
    mvmt_dt = mvmt_dt_ent.get()
    mem_id = membrane_id_combobox.get()
    act = act_ent.get()
    loc = loc_ent.get()
    pos = pos_ent.get()
    com = com_ent.get()
    integ = integ_ent.get()
    perm = perm_ent.get()
    tubes = tubes_ent.get()
    rep_ref = rep_ref_ent.get()
    pre = pre_ent.get()
    post = post_ent.get()

    if mvmt_dt and mem_id and act and loc:
        try:
            conn = sqlite3.connect(db_file_path)
            cursor = conn.cursor()
            cursor.execute( '''INSERT INTO movements (Datetime, Membrane_ID, Action, 
                       Location, Position, Comment, Integrity, Permeability, 
                       Tubes_sealed, Report_ref, Pre_clean_notes, Post_clean_notes) VALUES( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                       (mvmt_dt, mem_id, act, loc, pos, com, integ, perm, tubes, rep_ref, pre, post))
            conn.commit()
            conn.close()
            clear_entries()
            display_ufs()
        
        except sqlite3.Error as e:
            messagebox.showerror("Database Error", f"Error adding movement: {e}")
    
    else:
        messagebox.showerror("Error", "Please fill in all fields.")

# Function to clear the entry fields
def clear_entries():
    mvmt_dt_ent.delete(0, tk.END)
    membrane_id_combobox.delete(0, tk.END)
    act_ent.delete(0, tk.END)
    loc_ent.delete(0,tk.END)
    pos_ent.delete(0,tk.END)
    com_ent.delete(0,tk.END)
    integ_ent.delete(0,tk.END)
    perm_ent.delete(0,tk.END)
    tubes_ent.delete(0,tk.END)
    rep_ref_ent.delete(0,tk.END)
    pre_ent.delete(0,tk.END)
    post_ent.delete(0,tk.END)

# Function to populate the membrane IDs in the dropdown menu
def populate_membrane_ids():
    try:
        conn = sqlite3.connect(db_file_path)
        cursor = conn.cursor()

        # Fetch all unique membrane IDs from the 'movements' table
        cursor.execute('SELECT DISTINCT Membrane_ID FROM movements')
        membrane_ids = [row[0] for row in cursor.fetchall()]
        
        conn.close()

        # Clear any previous values in the Combobox
        membrane_id_combobox['values'] = ()
        membrane_id_combobox['values'] = tuple(membrane_ids)
    
    except sqlite3.Error as e:
        messagebox.showerror("Error", f"Database Error: {e}")

# Function to handle the selection of a membrane ID from the Combobox
def select_membrane_id(event):
    selected_membrane_id = membrane_id_combobox.get()
    # You can use the selected_membrane_id as needed

# Function to add new membranes to the database
def add_membrane():
    membrane_id = mem_ent.get()
    purchase_date = pur_ent.get()

    if membrane_id and purchase_date:
        try:
            conn = sqlite3.connect(db_file_path)
            cursor = conn.cursor()
            cursor.execute( '''INSERT INTO purchases (Membrane_ID, purchase_datetime) VALUES( ?, ?)''',
                       (membrane_id, purchase_date ))
            conn.commit()
            conn.close()
            display_mem_purchase()

        except sqlite3.Error as e:
            messagebox.showerror("Database Error", f"Error adding movement: {e}")
    
    else:
        messagebox.showerror("Error", "Please fill in all fields.")

def display_mem_purchase():
    conn = sqlite3.connect(db_file_path)
    cursor = conn.cursor()
    cursor.execute('''SELECT * from membranes;
                   ''')
    books = cursor.fetchall()
    conn.close()

    # Clear the listbox
    mem_listbox.delete()
    # print(books)
    # Insert books into the listbox
    for book in books:
        if book[2] !="xx":
            age = round((dt.datetime.now() - dt.datetime.strptime(book[2],"%Y-%m-%d %H:%M:%S")).days/365, 2)
        else:
            age = "Unkown"
        mem_listbox.insert('',tk.END,values=(book[1], str(age), book[2] ))

def data_entry(frame, name):
    x = ttk.Label(frame, text = name)
    x.pack()
    y = ttk.Entry(frame)
    y.pack()
    return y


# driver code
if __name__=="__main__":
    # Create the main window
    root = ThemedTk(theme='arc')
    root.title("UF Management")
    
    # ttk.Style().theme_use('Arc')

    tabControl = ttk.Notebook(root)
    tab1 = ttk.Frame(tabControl)
    # tab1.Style().theme_use('arc')
    tab2 = ttk.Frame(tabControl)

    tabControl.add(tab1, text = 'Add Movements')
    tabControl.add(tab2, text = 'Add Membranes')
    tabControl.pack(expand=1,fill="both")

    mvment_data_entry = ttk.Frame(master=tab1)
    pos_display = ttk.Frame(master=tab1)
    mvment_data_entry.pack(side=tk.LEFT, fill=tk.Y)
    pos_display.pack(side=tk.LEFT, fill=tk.BOTH)

    # Create labels and entry fields
    
    mvmt_dt_ent = data_entry(mvment_data_entry, name="Datetime(YYYY-mm-dd):")
    # Create a Combobox for selecting a membrane ID
    membrane_id_combobox = ttk.Combobox(mvment_data_entry)
    membrane_id_combobox .pack()   
    membrane_id_combobox.set("Select Membrane ID")  # Default text in the Combobox
    membrane_id_combobox.bind("<<ComboboxSelected>>", select_membrane_id)  # Bind event handler

    act_ent = data_entry(mvment_data_entry, name="Action:")
    loc_ent = data_entry(mvment_data_entry, name="Location:")    
    pos_ent = data_entry(mvment_data_entry, name="Position:")
    com_ent = data_entry(mvment_data_entry, name="Comments:")
    integ_ent = data_entry(mvment_data_entry, name="Ingegrity:")
    perm_ent = data_entry(mvment_data_entry, name="Permeability:")
    tubes_ent = data_entry(mvment_data_entry, name= "Sealed tubes:")
    rep_ref_ent = data_entry(mvment_data_entry, name="Report Referance:")
    pre_ent = data_entry(mvment_data_entry, name="Pre-clean notes:")
    post_ent = data_entry(mvment_data_entry, name="Post-clean notes")

    # Populate the Combobox with membrane IDs
    populate_membrane_ids()

    # Create buttons
    add_button = ttk.Button(mvment_data_entry, text="Add Mvmnt", command=add_mvmt)
    add_button.pack()

    clear_button = ttk.Button(mvment_data_entry, text="Clear Entries", command=clear_entries)
    clear_button.pack()   
    
    # Create a treeview to display books
    uf_cols = ('Location', 'Membrane_ID', 'Position', 'time_in_location','Age' )
    uf_listbox = ttk.Treeview(pos_display, columns= uf_cols, show='headings')
    uf_listbox.heading('Location', text='Location')
    uf_listbox.heading('Membrane_ID', text='Membrane ID')
    uf_listbox.heading('Position', text='Position')
    uf_listbox.heading('time_in_location', text='Time in Location (days)')
    uf_listbox.heading('Age', text='Membrane Age(Years)')
    
    uf_listbox.pack(fill=tk.BOTH, expand=True)

   # create frames for tab2 
    mem_data_entry = ttk.Frame(master=tab2)
    mem_data_entry.pack(side=tk.LEFT, fill=tk.Y)
    mem_display = ttk.Frame(master=tab2)
    mem_display.pack(side=tk.LEFT, fill=tk.BOTH )
    mem_ent = data_entry(mem_data_entry, "Membrane_ID:")
    pur_ent = data_entry(mem_data_entry, "DOM:")

    mem_button = ttk.Button(mem_data_entry, text="Add Membrane", command=add_membrane)
    mem_button.pack()

    mem_cols = ("Membrane_ID", "Age", "DOM")
    mem_listbox = ttk.Treeview(mem_display, columns = mem_cols, show="headings")
    mem_listbox.heading("Membrane_ID", text = 'Membrane ID')
    mem_listbox.heading("DOM", text = 'Date of Manufacture')
    mem_listbox.heading("Age", text = 'Age (Years)')
    mem_listbox.pack(fill=tk.BOTH, expand=True)

    # Display existing books in the listbox
    display_ufs()
    display_mem_purchase()

    root.mainloop()