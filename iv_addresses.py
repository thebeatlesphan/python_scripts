from bs4 import BeautifulSoup
import requests

with open("yellow.html") as fp:
	soup = BeautifulSoup(fp, "html.parser")

	print(soup.prettify())