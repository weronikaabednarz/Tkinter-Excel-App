import tkinter as tk
from tkinter import ttk

window = tk.Tk()

style = ttk.Style(window)       # enable to apply a theme to app
window.tk.call("source", "forest-light.tcl")
window.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")
#style.theme_use("forest-light")



window.mainloop()