#regular expression library
import re

#coordinates.txt comes from https://www.infoplease.com/us/geography/latitude-and-longitude-us-and-canadian-cities
lines = []
with open("coordinates.txt", "r") as f:
	lines = f.readlines()

new_file = open("cities.txt", "a")

for line in lines:
	#regular expressions to grab desired results
	city = re.search('\w+(?=,)', line)
	coordinates = re.findall('\t\d+\t', line)
	new_file.write("\"{}\": \"{} {}\",\n".format(city[0],coordinates[0].strip(),coordinates[1].strip()) )
new_file.close()
