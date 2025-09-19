# testing.py 20250917

import csv
from openpyxl import Workbook

CSV_File = "data.csv"

data = None
with open(CSV_File, newline='') as csvfile:
	csvdata = csv.reader(csvfile, delimiter=' ', quotechar='|')
	data = [row for row in csvdata]
#	for row in csvdata:
#		print(row)

wb = Workbook()


