import tkinter as tk
from tkinter import ttk
import openpyxl

options_list = ["Programming", "Management", "Marketing", "Data analysis", "Computer graphics", "Web design", "Accounting", "International trade", "Employment law"]

def load_data():
    path = "C:\\Users\\weron\\Desktop\\projekty\\Tkinter Excel App\\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    print(list_values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text = col_name)
    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values = value_tuple)

def insert_row():
    name = name_entry.get()
    age = int(age_spinbox.get())
    interests = interests_combobox.get()
    employment_status = "Employed" if status.get() else "Unemployed"
    print(name, age, interests, employment_status)

    # insert row into Excel sheet
    path = "C:\\Users\\weron\\Desktop\\projekty\\Tkinter Excel App\\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, interests, employment_status]
    sheet.append(row_values)
    workbook.save(path)

    # insert row into treeview
    treeview.insert('', tk.END, values = row_values)

    # clear values
    name_entry.delete(0, "end")
    name_entry.insert(0, 'Name')
    age_spinbox.delete(0, "end")
    age_spinbox.insert(0, "Age")
    interests_combobox.set(options_list[0])
    employed_checkbutton.state(["!selected"])

def toggle_mode():
    if mode_switch.instate(["selected"]):       # selected - press on it once
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

window = tk.Tk()
window.title("Excel Application")

style = ttk.Style(window)       # enable to apply a theme to app
window.tk.call("source", "forest-light.tcl")
window.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

frame = ttk.Frame(window)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text = "Insert Row")
widgets_frame.grid(row = 0, column = 0, padx = 20, pady = 10)

name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Name")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))    # focus in - press on entry
name_entry.grid(row = 0, column = 0, padx = 5, pady = (0,5), sticky = "ew")     # 0 pixels from the top edge and 5 pixels from the bottom edge

age_spinbox = ttk.Spinbox(widgets_frame, from_ = 18, to = 100)
age_spinbox.insert(0, "Age")
age_spinbox.bind("<FocusIn>", lambda e: age_spinbox.delete('0', 'end'))
age_spinbox.grid(row = 1, column = 0, padx = 5, pady = 5, sticky = "ew")

interests_combobox = ttk.Combobox(widgets_frame, values = options_list)
interests_combobox.current(0)
interests_combobox.grid(row = 2, column = 0, padx = 5, pady = 5, sticky = "ew")

status = tk.BooleanVar()
employed_checkbutton = ttk.Checkbutton(widgets_frame, text = "Employed", variable = status)
employed_checkbutton.grid(row = 3, column = 0, padx = 5, pady = 5, sticky = "w")

button = ttk.Button(widgets_frame, text = "Insert", command = insert_row)
button.grid(row = 4, column = 0, sticky = "news")

separator = ttk.Separator(widgets_frame)
separator.grid(row = 5, column = 0, padx = [20,10], pady = 10, sticky = "ew")   # 20 pixels on the left side and 10 pixels on the right side

mode_switch = ttk.Checkbutton(widgets_frame, text = "Mode", style = "Switch", command = toggle_mode)   # style - defines type of checkbutton
mode_switch.grid(row = 6, column = 0, padx = 5, pady = 10, sticky = "ew")

treeFrame = ttk.Frame(frame)
treeFrame.grid(row = 0, column = 1, pady = 10)
#----------------tree scroll----------------------------------------------------------------------------------------------------------------
treeScroll = ttk.Scrollbar(treeFrame)   # scroll button, right side, full axis y
treeScroll.pack(side = "right", fill = "y")

cols = ("Name", "Age", "Interests", "Employment")
treeview = ttk.Treeview(treeFrame, show = "headings",
                        yscrollcommand = treeScroll.set, columns = cols, height = 11)
treeview.column("Name", width = 140)
treeview.column("Age", width = 50)
treeview.column("Interests", width = 140)
treeview.column("Employment", width = 100)
treeview.pack()
treeScroll.config(command = treeview.yview)
load_data()

window.mainloop()