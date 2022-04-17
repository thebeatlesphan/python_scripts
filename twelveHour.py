import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

#Read raw 'Incidents over 12 hours' report
#Then create a df of Micheline's group and desired columns
df = pd.read_excel("1.xlsx", usecols="N,A,D,B,G,L")
group = ["Andy Phan", "Martel Perrin"]
only_andy = df[df["Assigned to"].isin(group)]
only_andy["Updates / Comments / NOTES"] = " "
cols = only_andy.columns.tolist()
myorder = [5, 0, 2, 1, 3, 4, 6]
new_df = only_andy[cols].sort_values(by=["Assigned to"], ascending=True)
new_df['Created'] = new_df['Created'].dt.strftime('%Y-%m-%d')

#write df to excel and format the file to Michelines S: folder
with pd.ExcelWriter("S:/NorcrossTAMS/MPhillips/12hour/new.xlsx") as writer:
	new_df.to_excel(writer, index=False, sheet_name="andy")
	worksheet = writer.sheets["andy"]
	workbook = writer.book
	worksheet.set_default_row(45, hide_unused_rows=True)
	worksheet.set_column('H:XFD', None, None, {'hidden': True})
	worksheet.set_row(0, 14.4)
	format = workbook.add_format({'text_wrap': True, 'bottom':1, 'top':1, 'left':1, 'right':1})
	format.set_align('center')
	format.set_align('vcenter')

	writer.sheets["andy"].set_column("A:A", 15, format)
	writer.sheets["andy"].set_column("B:B", 15, format)
	writer.sheets["andy"].set_column("C:C", 20, format)
	writer.sheets["andy"].set_column("D:D", 40, format)
	writer.sheets["andy"].set_column("E:E", 40, format)
	writer.sheets["andy"].set_column("F:F", 10, format)
	writer.sheets["andy"].set_column("G:G", 40, format)
	writer.save()