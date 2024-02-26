import tkinter as tk
from tkinter import ttk

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

interests_combobox = ttk.Combobox(widgets_frame, values = ["Programming", "Management", "Marketing",
                                                           "Data analysis", "Computer graphics", "Web design",
                                                           "Accounting", "International trade""Employment law"])
interests_combobox.current(0)
interests_combobox.grid(row = 2, column = 0, padx = 5, pady = 5, sticky = "ew")

a = tk.BooleanVar()
employed_checkbutton = ttk.Checkbutton(widgets_frame, text = "Employed", variable = a)
employed_checkbutton.grid(row = 3, column = 0, padx = 5, pady = 5, sticky = "w")

button = ttk.Button(widgets_frame, text = "Insert")
button.grid(row = 4, column = 0, sticky = "news")

separator = ttk.Separator(widgets_frame)
separator.grid(row = 5, column = 0, padx = [20,10], pady = 10, sticky = "ew")   # 20 pixels on the left side and 10 pixels on the right side

mode_switch = ttk.Checkbutton(widgets_frame, text = "Mode", style = "Switch", command = toggle_mode)   # style - defines type of checkbutton
mode_switch.grid(row = 6, column = 0, padx = 5, pady = 10, sticky = "ew")


window.mainloop()