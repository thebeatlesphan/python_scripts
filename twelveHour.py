# Name: twelveHour.py
# Details: Program to transform the daily 12 hour incident RPT into Micheline's own report.
# Author: Andy Phan
# Created 2022-5-06
# Compile the program to an exe via auto-python-to-exe library

'''
June 10th, 2022
- Completed incorporating an INI file for group_names
- Completed new UI

May 21st, 2022
- Implemented Treeview and configparser

May 7th, 2022
- Removed "Justin Quesada" and "Fernando Robles" from group variable
- Added condition to format date "Created" column to change font color if >= 25 days ago
'''

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import warnings
import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from datetime import date, timedelta
import configparser

#Global variables
filepath = ""
df = ""
today = str(date.today() - timedelta(25))
config_file = ""
group_names = []

if exists(file_path):
	config_file = configparser.ConfigParser()
	config_file.read(file_path)
	for key in config_file["Name"]:
		group_names.append(config_file["Name"][key])

else:
	os.mkdir(os.path.expanduser('~/Documents/twelveHour'))
	config_file = configparser.ConfigParser()
	config_file.add_section("Name")
	with open(file_path, 'w') as config:
		config_file.write(config)

def open_file():
	global df

	#Grab file and display filepath
	select_file = askopenfilename()
	file["text"] = select_file
	filepath = select_file
	btn_save["state"] = "normal"

	#Create desired df from filepath
	with warnings.catch_warnings(record=True):
		df = pd.read_excel(filepath, usecols="N,A,D,B,G,L", engine="openpyxl")
		group = ["Andy Phan", "Martel Perrin", "Andrew Plourde", "Antoine Adderley", "Cedell Okegbenro", "Christopher De Guzman", "Clarence Spearman", "James Thompson1", "Maya Mattison", "Preston Hook", "Ron Wells", "Tim Stoklas", "Zachary Brown"]
		desired_team = df[df["Assigned to"].isin(group)]
		desired_team["Updates / Comments / NOTES"] = ""
		cols = desired_team.columns.tolist()
		myorder = [5,0,2,1,3,4,6]
		cols = [cols[i] for i in myorder]
		df = desired_team[cols].sort_values(by=["Assigned to"], ascending=True)
		df["Created"] = df["Created"].dt.strftime('%Y-%m-%d')

	status_indicator.configure(background="yellow")
	status_indicator["text"] = 'File Loaded'

def save_file():
	global df
	save_as = text.get()
	to_michelines = "S:/NorcrossTAMS/MPhillips/12hour/"+save_as+".xlsx"

	#Write df to excel and format it
	with pd.ExcelWriter(to_michelines) as writer:
		new_df.to_excel(writer, index=False, sheet_name="andy")
		worksheet = writer.sheets["andy"]
		workbook = writer.book
		worksheet.set_default_row(45, hide_unused_rows=True)
		worksheet.set_column('H:XFD', None, None, {'hidden': True})
		worksheet.set_row(0, 14.4)
		format = workbook.add_format({'text_wrap': True, 'bottom':1, 'top':1, 'left':1, 'right':1})
		format.set_align('center')
		format.set_align('vcenter')
		red_format = workbook.add_format({'font_color': 'red'})
		worksheet.conditional_format("C2:C500", {'type': 'cell', 'critieria': '<=', 'value': '$H$2', 'format': red_format})
		writer.sheets["andy"].set_column("A:A", 15, format)
		writer.sheets["andy"].set_column("B:B", 15, format)
		writer.sheets["andy"].set_column("C:C", 20, format)
		writer.sheets["andy"].set_column("D:D", 40, format)
		writer.sheets["andy"].set_column("E:E", 40, format)
		writer.sheets["andy"].set_column("F:F", 10, format)
		writer.sheets["andy"].set_column("G:G", 40, format)
		
	status_indicator.configure(background="green")
	status_indicator["text"] = 'File Saved'

def add_name():
	global group_names

	group_names.append(add_name_entry.get())
	tree.insert('', tk.END, values=(add_name_entry.get(),))
	config_file["Name"][add_name_entry.get()] = add_name_entry.get()
	with open(file_path, 'w+') as config:
		config_file.write(config)
	add_name_entry.delete(0,100)

def delete_name():
	global group_names

	selected_item = tree.selection()
	for item in selected_item:
		config_file.remove_option("Name", tree.item(item)['values'][0])
		group_names.remove(tree.item[item]['values'][0])
		tree.delete(item)
	with open(file_path, 'w+') as config:
		config_file.write(config)

# GUI for program
window = tk.Tk()

window.title("Micheline's Little Helper")
window.geometry('330x140')
window.resizable(width=False, height=False)
frame = tk.Frame(master=window)

btn_open = tk.Button(master=frame, text="Open", command=open_file)
btn_save = tk.Button(master=frame, text="Save As...", command=save_file)

status = tk.Label(window, text="File Status").grid(row=0,column=0)
status_indicator = tk.Label(window, text="No File Loaded").grid(row=1,column=0,rowspan=2)

add_name_entry = tk.Entry(window).grid(row=4,column=0)

btn_add = tk.Button(window, text="Add Agent", command=add_name).grid(row=5,column=0)
btn_delete = tk.Button(window, text="Delete Agent", command=delete_name).grid(row=5, column=1)

# Employee Treeview List
tree = ttk.Treeview(window, height=4, columns=("Name"), show='headings')
tree.grid(row=0, column=1, rowspan=5, sticky=tk.E)
tree.heading("Name", text="Name")

# Dropdown Menu
menubar = Menu(window)
window.config(menu=menubar)
file_menu = Menu(menubar, tearoff=False)
file_menu.add_command(label='Open', command=open_file)
file_menu.add_command(label='Save as', command=save_file)
file_menu.add_command(label='Exit', command=window.destroy)
menubar.add_cascade(label="File", menu=file_menu)

def update_tree():
	for name in group_names:
		tree.insert('', tk.END, values=(name,))

update_tree()
window.mainloop()