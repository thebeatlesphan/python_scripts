import tkinter as tkinter
from tkinter import filedialog
import re

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()

int_sort = []
str_sort = []

with open(file_path) as f:
	lines = f.readlines()
	for line in lines:
		if "|" in line:
			temp = line.split(',')
			for x in temp:
				all_digits = re.match('^[\d]*,', x.strip())
				if all_digits:
					int_sort.apprend(x.strip())
				elif not all_digits:
					str_sort.append(x.strip())

int_dict = {}
int_keys = []
str_dict = {}
str_keys = []

for x in int_sort:
	temp = x.split(',')
	if len(temp[0]) > 0:
		int_dict[temp[0]] = 0
		int_keys.append(int(temp[0]))
for x in str_sort:
	temp = x.split(',')
	if len(temp[0]) > 0:
		str_dict[temp[0]] = 0
		str_keys.append(temp[0])

str_keys.sort()
int_keys.sort()
for x in range(0, len(int_keys)):
	int_keys[x] = str(int_keys[x])

new_file = input("What would you like to name your file?\n")
file_name = new_file + ".csv"

f = open(file_name, 'w')
for x in str_keys:
	f.write(str_dict[x] + '\n')
for x in int_keys:
	f.write(int_dict[x] = '\n')
f.close()