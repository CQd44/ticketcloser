from easygui import *
import toml
import psycopg2
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import csv

CONFIG = toml.load("./config.toml") #load variables from toml file
COLUMNS: list[str] = ['id', 'name', 'clinic', 'issue', 'entrydate', 'open', 'remarks']
data: list = []
SELECTED_ROW: int = 0
OPEN = 'ALL'

def set_open_true():
    global OPEN
    OPEN = True
    update_data()

def set_open_false():
    global OPEN
    OPEN = False
    update_data()

def set_all():
    global OPEN
    OPEN = 'ALL'
    update_data()

def select_row(event):
    global SELECTED_ROW
    selected_item = table.selection()[0]  # Get the selected item's ID
    if selected_item:
        row_data = table.item(selected_item)['values'] # Get data of the selected row
        SELECTED_ROW = row_data[0]

def delete_row():
    if SELECTED_ROW == 0:
        msgbox(title = 'No Row Selected', msg = 'No row selected!', ok_button = 'OK')
    else:
        con = psycopg2.connect(f'dbname = {CONFIG['credentials']['dbname']} user = {CONFIG['credentials']['username']} password = {CONFIG['credentials']['password']}')
        cur = con.cursor()
        SQL = 'SELECT * FROM responses WHERE id = %s;'
        DATA = SELECTED_ROW
        cur.execute(SQL, (DATA,))
        result = cur.fetchone()[0]
        delete = ynbox(msg = "Really delete this row?", 
                       title = "Row Deletion", 
                       choices = ("[<F1>]Yes", "[<F2>]No"), 
                       default_choice = '[<F2>]No', 
                       cancel_choice = '[<F2>]No')
        if delete == False:
            pass
        if delete == True:
            SQL = 'DELETE FROM responses WHERE id = %s;'
            DATA = result
            cur.execute(SQL, (DATA,))
            cur.close()
            con.commit()
            msgbox(title = 'Row Deleted',
                   msg = f'Deleted row with Ticket ID {result}')
            update_data()

def update_data():
    data = [] #if you don't reinitialize this it just adds duplicate entries every time
    refresh_data(table, data, OPEN)

def refresh_data(table, data, OPEN):
    global SELECTED_ROW
    SELECTED_ROW = 0
    for item in table.get_children():
        table.delete(item)
    con = psycopg2.connect(f'dbname = {CONFIG['credentials']['dbname']} user = {CONFIG['credentials']['username']} password = {CONFIG['credentials']['password']}')
    cur = con.cursor()
    DATA = OPEN
    if OPEN != 'ALL':
        SQL = 'SELECT * FROM responses WHERE open is %s ORDER BY id ASC;'        
        cur.execute(SQL, (DATA,))
    if OPEN == 'ALL':
        SQL = 'SELECT * FROM responses ORDER by id ASC;'
        cur.execute(SQL)    
    rows = cur.fetchall()
    cur.close()
    con.close()

    for row in rows:
        subresult = []
        for i in range(len(COLUMNS)):
            subresult.append(row[i])
        data.append(subresult)

    for row in data:
        table.insert("", tk.END, values=row)

def search():
    global table
    global data
    global SELECTED_ROW

    SELECTED_ROW = 0
    data = []

    search_text = enterbox(title = 'Ticket Search', msg = 'Enter search criteria (all parts of all tickets will be searched):', default = '')
    search_text = '%' + search_text + '%'
    for item in table.get_children():
        table.delete(item)
    con = psycopg2.connect(f'dbname = {CONFIG['credentials']['dbname']} user = {CONFIG['credentials']['username']} password = {CONFIG['credentials']['password']}')
    cur = con.cursor()
    DATA = (search_text, search_text, search_text, search_text) #clumsy but works
    cur.execute("""SELECT * FROM responses WHERE 
                (issue LIKE %s OR
                name LIKE %s OR
                clinic LIKE %s OR
                remarks LIKE %s)
                ORDER BY id ASC;""", DATA)
    rows = cur.fetchall()
 
    cur.close()
    con.close()

    for row in rows:
        subresult = []
        for i in range(len(COLUMNS)):
            subresult.append(row[i])
        data.append(subresult)

    for row in data:
        table.insert("", tk.END, values=row)

def close_ticket():
    global SELECTED_ROW
    if SELECTED_ROW == 0:
        msgbox(title = 'No Ticket Selected', 
               msg = 'You need to select a ticket before you can close it!')    
    else:
        con = psycopg2.connect(f'dbname = {CONFIG['credentials']['dbname']} user = {CONFIG['credentials']['username']} password = {CONFIG['credentials']['password']}')
        cur = con.cursor()
        cur.execute('SELECT open FROM responses WHERE id = %s', (SELECTED_ROW,))
        result = cur.fetchone()
        if result[0] == False:
            msgbox(title = 'Ticket Already Closed', 
                   msg = "Can't close a ticket that isn't open!")
        if result[0] == True:
            close_question: bool = ynbox(title = "Close Selected Ticket", 
                                    msg = "Do you want to close this ticket?", 
                                    choices = ("[<F1>]Yes", "[<F2>]No"), 
                                    default_choice = '[<F2>]No', 
                                    cancel_choice = '[<F2>]No')
            if close_question == False:
                pass
            if close_question == True:
                closing_remarks: str = textbox(title = "Ticket's Closing Remarks", 
                                            msg = "Please enter your closing remarks for this ticket.",
                                            text = "None (no remarks given)")

                SQL = 'UPDATE responses SET open = False, remarks = %s WHERE id = %s;'
                DATA = (closing_remarks, SELECTED_ROW)
                cur.execute(SQL, DATA)        
                cur.close()
                con.commit()
                msgbox(title = 'Ticket Closed', msg = 'The selected ticket has been successfully closed! :)')
                update_data()

def generate_report():
    REPORT_COLUMNS: list[str] = ['id', 'name', 'clinic', 'issue', 'entrydate', 'open', 'remarks', 'ip', 'hostname']
    RESULTS: list[any] = []

    con = psycopg2.connect(f'dbname = {CONFIG['credentials']['dbname']} user = {CONFIG['credentials']['username']} password = {CONFIG['credentials']['password']}')
    cur = con.cursor()

    cur.execute("SELECT * FROM responses")
    rows = cur.fetchall()

    cur.close()
    
    output = filesavebox(title = 'Select filename and location to save report',
                          msg = 'Filename: ', 
                          default='report.csv', 
                          filetypes = '*.csv')

    if output:
        for row in rows:
            subresult = []
            for i in range(len(REPORT_COLUMNS)):
                subresult.append(row[i])
            RESULTS.append(subresult)

        with open(output, 'w', newline = '') as report:
            writer = csv.writer(report)
            writer.writerow(REPORT_COLUMNS)
            for row in RESULTS:
                writer.writerow(row)

root = tk.Tk()
root.title("Call Center Tickets")

main_frame = ttk.Frame(root)
main_frame.pack(pady=10)

table = ttk.Treeview(root, columns=COLUMNS, show="headings", height=20)
table.pack(side=tk.LEFT, fill=tk.BOTH)

# Define column headings
for column in COLUMNS:
    table.heading(f'{column}', text=f'{column}')

scrollbar = ttk.Scrollbar(root, orient=tk.VERTICAL, command=table.yview)
scrollbar.pack(side=tk.LEFT, fill=tk.Y)

table.configure(yscrollcommand=scrollbar.set)

# Button Frame
button_frame = ttk.Frame(root)
button_frame.pack(side=tk.TOP, pady = 10)

# Actual Buttons
show_row = ttk.Button(button_frame, text="Search Tickets", command=search)
show_row.pack(padx = 5, pady = 5)

refresh_button = ttk.Button(button_frame, text="Refresh Database", command=update_data)
refresh_button.pack(padx = 5, pady = 5)

delete_button = ttk.Button(button_frame, text="Delete Row", command=delete_row)
delete_button.pack(padx = 5, pady = 5)

view_open_tickets = ttk.Button(button_frame, text = 'Show Open Tickets', command = set_open_true)
view_open_tickets.pack(padx = 5, pady = 5)

view_closed_tickets = ttk.Button(button_frame, text = 'Show Closed Tickets', command = set_open_false)
view_closed_tickets.pack(padx = 5, pady = 5)

view_all_tickets = ttk.Button(button_frame, text = 'Show All Tickets', command = set_all)
view_all_tickets.pack(padx = 5, pady = 5)

close_selected_ticket = ttk.Button(button_frame, text = 'Close Ticket', command = close_ticket)
close_selected_ticket.pack(padx = 5, pady = 5)

report = ttk.Button(button_frame, text = 'Generate Report of Current View', command = generate_report)
report.pack(padx = 5, pady = 5)

#Load DHR logo
image = Image.open("dhr-logo.jpg") # DHR logo
resized_image = image.resize((400, 110))
photo = ImageTk.PhotoImage(resized_image)
image_label = tk.Label(main_frame, image=photo)
image_label.image = photo # Keep a reference to prevent garbage collection
image_label.pack(side = tk.BOTTOM)

# Initial table population
refresh_data(table, data, OPEN)

table.bind("<ButtonRelease-1>", select_row)  # Call select_row on mouse click

table.pack()
root.mainloop()