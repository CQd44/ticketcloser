from easygui import *
import toml
import psycopg2
import psycopg2.sql as sql
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import csv

CONFIG = toml.load("./config.toml") # load variables from toml file. 

class OpenState: # Avoids using global variables
    OPEN: any = 'ALL' # Keeps track of whether you're looking at open, closed, or all tickets
    SELECTED_ROW: int = 0 # I shouldn't need to explain this. 
    data: list = [] # list that holds the data the table / treeview is comprised of
    view_cols: list[str] = ['id', 'name', 'clinic', 'issue', 'entrydate', 'open', 'remarks'] # default view
    root: any = None # main window object
    table: any # main table that shows data
    scrollbar: any # vertical scrollbar
    hscrollbar: any # horizontal scrollbar
    main_frame: any # where things get stuffed in the main window object
    button_frame: any # where the buttons go
    image = Image.open("dhr-logo.jpg") # the actual image
    image_label: any # where the logo goes

def set_open_true():
    OpenState.OPEN = True
    update_data()

def set_open_false():
    OpenState.OPEN = False
    update_data()

def set_all():
    OpenState.OPEN = 'ALL'
    update_data()

def select_row(event):
    try:
        selected_item = OpenState.table.selection()[0]  # Get the selected item's ID
        if selected_item:
            row_data = OpenState.table.item(selected_item)['values'] # Get data of the selected row
            OpenState.SELECTED_ROW = row_data[0]
    except:
        OpenState.SELECTED_ROW = 0

def delete_row():
    if OpenState.SELECTED_ROW == 0:
        msgbox(title = 'No Row Selected', msg = 'No row selected!', ok_button = 'OK')
    else:
        con = psycopg2.connect(f'host = {CONFIG['credentials']['host']} dbname = {CONFIG['credentials']['dbname']} user = {CONFIG['credentials']['username']} password = {CONFIG['credentials']['password']}')
        cur = con.cursor()
        SQL = 'SELECT * FROM responses WHERE id = %s;'
        DATA = OpenState.SELECTED_ROW
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
    OpenState.data = [] #if you don't reinitialize this it just adds duplicate entries every time
    refresh_data(OpenState.table, OpenState.data, OpenState.OPEN)

def refresh_data(table, data, open):
    OpenState.SELECTED_ROW = 0
    for item in table.get_children():
        OpenState.table.delete(item)
    con = psycopg2.connect(f'host = {CONFIG['credentials']['host']} dbname = {CONFIG['credentials']['dbname']} user = {CONFIG['credentials']['username']} password = {CONFIG['credentials']['password']}')
    cur = con.cursor()
    DATA = (open,)
    if open != 'ALL':
        query = sql.SQL("SELECT {} FROM responses WHERE open IS %s ORDER BY 1 ASC;").format(
            sql.SQL(",").join(sql.Identifier(column) for column in OpenState.view_cols)
        )
        cur.execute(query.as_string(con), DATA)
    else:
        query = sql.SQL("SELECT {} FROM responses ORDER BY id ASC;").format(
            sql.SQL(",").join(sql.Identifier(column) for column in OpenState.view_cols)
        )
        cur.execute(query.as_string(con))
    rows = cur.fetchall()
    cur.close()
    con.close()

    for row in rows:
        subresult = []
        for i in range(len(OpenState.view_cols)):
            subresult.append(row[i])
        data.append(subresult)

    for row in data:
        OpenState.table.insert("", tk.END, values=row)

def sort_treeview(tree, col, descending):
    data = [(OpenState.table.set(item, col), item) for item in OpenState.table.get_children('')]
    try:
        data.sort(key = lambda t: int(t[0]), reverse=descending) #special case for ints (like ID)
    except ValueError:
        data.sort(reverse=descending) #all other columns, since you can't int strings :) 
    
    for index, (val, item) in enumerate(data):
        tree.move(item, '', index)
    OpenState.table.heading(col, command = lambda: sort_treeview(tree, col, not descending))

def search():
    OpenState.SELECTED_ROW = 0
    OpenState.data = []
    search_text = enterbox(title = 'Ticket Search', msg = 'Enter search criteria (all parts of all tickets will be searched):', default = '')
    if search_text != None:
        search_text = '%' + search_text + '%'
        for item in OpenState.table.get_children():
            OpenState.table.delete(item)
        con = psycopg2.connect(f'host = {CONFIG['credentials']['host']} dbname = {CONFIG['credentials']['dbname']} user = {CONFIG['credentials']['username']} password = {CONFIG['credentials']['password']}')
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
            for i in range(len(OpenState.view_cols)):
                subresult.append(row[i])
            OpenState.data.append(subresult)

        for row in OpenState.data:
            OpenState.table.insert("", tk.END, values=row)
    else:
        pass

def close_ticket():
    if OpenState.SELECTED_ROW == 0:
        msgbox(title = 'No Ticket Selected', 
               msg = 'You need to select a ticket before you can close it!')    
    else:
        con = psycopg2.connect(f'host = {CONFIG['credentials']['host']} dbname = {CONFIG['credentials']['dbname']} user = {CONFIG['credentials']['username']} password = {CONFIG['credentials']['password']}')
        cur = con.cursor()
        cur.execute('SELECT open FROM responses WHERE id = %s', (OpenState.SELECTED_ROW,))
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
                                            msg = "Please enter your closing remarks for this ticket.")

                SQL = 'UPDATE responses SET open = False, remarks = %s WHERE id = %s;'
                DATA = (closing_remarks, OpenState.SELECTED_ROW)
                cur.execute(SQL, DATA)        
                cur.close()
                con.commit()
                msgbox(title = 'Ticket Closed', msg = f'The selected ticket (ID: {OpenState.SELECTED_ROW}) has been successfully closed! :)')
                update_data()

def select_view():
    OpenState.view_cols = multchoicebox(title = 'View Selector', 
                                        msg = 'Select the columns you wish to view',
                                        choices = ['id', 'name', 'clinic', 'issue', 'entrydate', 'open', 'remarks', 'ip_addr', 'hostname'],
                                        preselect = [0, 1, 2, 3, 4, 5, 6]
                                        )
    if OpenState.view_cols == None:
        msgbox(title = 'Error!',
               msg = 'At least one selection must be made!')
        select_view()
    else:
        OpenState.button_frame.destroy()
        OpenState.image_label.destroy()
        OpenState.table.destroy()
        OpenState.scrollbar.destroy()
        OpenState.main_frame.destroy()
        create_window()
        update_data()

def generate_report():
    RESULTS: list[any] = []

    con = psycopg2.connect(f'host = {CONFIG['credentials']['host']} dbname = {CONFIG['credentials']['dbname']} user = {CONFIG['credentials']['username']} password = {CONFIG['credentials']['password']}')
    cur = con.cursor()

    if OpenState.OPEN != 'ALL':
        query = sql.SQL("SELECT {} FROM responses WHERE open IS %s ORDER BY 1 ASC;").format(
            sql.SQL(",").join(sql.Identifier(column) for column in OpenState.view_cols)
        )
        cur.execute(query.as_string(con), (OpenState.OPEN,))
    else:
        query = sql.SQL("SELECT {} FROM responses ORDER BY id ASC;").format(
            sql.SQL(",").join(sql.Identifier(column) for column in OpenState.view_cols)
        )
        cur.execute(query.as_string(con))
    rows = cur.fetchall()

    cur.close()
    
    output = filesavebox(title = 'Select filename and location to save report',
                          msg = 'Filename: ', 
                          default='report.csv', 
                          filetypes = '*.csv')

    if output:
        for row in rows:
            subresult = []
            for i in range(len(OpenState.view_cols)):
                subresult.append(row[i])
            RESULTS.append(subresult)

        with open(output, 'w', newline = '') as report:
            writer = csv.writer(report)
            writer.writerow(OpenState.view_cols)
            for row in RESULTS:
                writer.writerow(row)

def create_window(): #assemble the actual window and buttons the user interacts with
    if not OpenState.root:
        OpenState.root = tk.Tk()
        OpenState.root.minsize(600, 475)
        OpenState.root.title("Call Center Tickets")

    OpenState.main_frame = ttk.Frame(OpenState.root)
    OpenState.main_frame.pack(pady=10)

    OpenState.table = ttk.Treeview(OpenState.root, columns=OpenState.view_cols, selectmode='browse', show="headings", height=20)
    OpenState.table.pack(side=tk.LEFT, fill=tk.BOTH, expand = True)

    # Define column headings
    for column in OpenState.view_cols:
        OpenState.table.heading(f'{column}', text=f'{column}', anchor = tk.CENTER, command = lambda c = f'{column}': sort_treeview(OpenState.table, c, False))
        OpenState.table.column(column, anchor=tk.CENTER, width = 80)
    OpenState.table.update()
    for column in OpenState.table['columns']:
        OpenState.table.column(column, width = 100, stretch = 0)

    OpenState.scrollbar = ttk.Scrollbar(OpenState.root, orient=tk.VERTICAL, command=OpenState.table.yview)
    OpenState.scrollbar.pack(side=tk.LEFT, fill=tk.Y)

    OpenState.hscrollbar = ttk.Scrollbar(OpenState.table, orient=tk.HORIZONTAL, command=OpenState.table.xview)
    OpenState.hscrollbar.pack(side=tk.BOTTOM, fill=tk.X)

    OpenState.table.configure(yscrollcommand=OpenState.scrollbar.set)
    OpenState.table.configure(xscrollcommand=OpenState.hscrollbar.set)

    # Button Frame
    OpenState.button_frame = ttk.Frame(OpenState.root)
    OpenState.button_frame.pack(side=tk.TOP, pady = 10)

    # Actual Buttons

    view_change = ttk.Button(OpenState.button_frame, text = 'Change View', command = select_view)
    view_change.pack(padx = 5, pady = 5)

    show_row = ttk.Button(OpenState.button_frame, text="Search Tickets", command=search)
    show_row.pack(padx = 5, pady = 5)

    refresh_button = ttk.Button(OpenState.button_frame, text="Refresh Database", command=update_data)
    refresh_button.pack(padx = 5, pady = 5)

    delete_button = ttk.Button(OpenState.button_frame, text="Delete Row", command=delete_row)
    delete_button.pack(padx = 5, pady = 5)

    view_open_tickets = ttk.Button(OpenState.button_frame, text = 'Show Open Tickets', command = set_open_true)
    view_open_tickets.pack(padx = 5, pady = 5)

    view_closed_tickets = ttk.Button(OpenState.button_frame, text = 'Show Closed Tickets', command = set_open_false)
    view_closed_tickets.pack(padx = 5, pady = 5)

    view_all_tickets = ttk.Button(OpenState.button_frame, text = 'Show All Tickets', command = set_all)
    view_all_tickets.pack(padx = 5, pady = 5)

    close_selected_ticket = ttk.Button(OpenState.button_frame, text = 'Close Ticket', command = close_ticket)
    close_selected_ticket.pack(padx = 5, pady = 5)

    report = ttk.Button(OpenState.button_frame, text = 'Generate Report of Current View', command = generate_report)
    report.pack(padx = 5, pady = 5)

    # Load DHR logo
    OpenState.image = Image.open("dhr-logo.jpg") # DHR logo
    resized_image = OpenState.image.resize((400, 110))
    photo = ImageTk.PhotoImage(resized_image)
    OpenState.image_label = tk.Label(OpenState.main_frame, image=photo)
    OpenState.image_label.image = photo # Keep a reference to prevent garbage collection
    OpenState.image_label.pack(side = tk.BOTTOM)

    # Initial table population
    refresh_data(OpenState.table, OpenState.data, OpenState.OPEN)

    OpenState.table.bind("<ButtonRelease-1>", select_row)  # Call select_row on mouse click

    OpenState.table.pack()

create_window()
OpenState.root.mainloop()