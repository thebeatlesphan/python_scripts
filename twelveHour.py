# Name: twelveHour.py
# Details: Program to transform the daily 12 hour incident RPT into Micheline's own report.
# Author: Andy Phan
# Created 2022-5-06
# Compile the program to an exe via auto-python-to-exe library

'''
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

#Global variables
filepath = ""
df = ""
today = str(date.today() - timedelta(25))

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
		
	file["text"] = 'DONE'

#GUI for program
window = tk.Tk()
window.title("12hour")

window.resizable(width=False, height=False)
frame = tk.Frame(master=window, width=250, height=150)
frame.pack()

description = tk.Label(text="Micheline's Little Helper", font=8).place(x=20, y=20)

btn_open = tk.Button(master=frame, text="Open", command=open_file).place(x=70, y=70)
btn_save = tk.Button(master=frame, text="Save As...", command=save_file)
btn_save.place(x=120, y=70)

text = tk.Entry(width=39)
text.place(x=4, y=110)

file = tk.Label(master=frame, text=filepath, wraplength=250)
file.place(x=0, y=130)

window.mainloop()