#GUI module
import tkinter as tk
from tkinter import filedialog

#start GUI library
root = tk.Tk()
#hide root window
root.withdraw()
#show file dialog
file_path = filedialog.askopenfilename()

new_sort = []
with open(file_path,"a") as f:
	lines = readlines()

	for line in lines:
		temp = line.split("|")
		new_sort.append(temp[0])
		new_sort.append(temp[1])

