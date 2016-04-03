import openpyxl
import webbrowser

# figure out a way to generate the url name (same name, but appended date)
workbookName = 'elearning.xlsx'

workbook = openpyxl.load_workbook(workbookName)
worksheet = workbook.get_sheet_by_name('Sheet1')
sheetValue = 1  # initialize
linkRow = 1
linkColumn = 5  # the column where the links are

# need to figure out a way to compare an empty str to None type

while sheetValue is not None:  # while there is data in that cell
	sheetValue = worksheet.cell(row=linkRow, column=linkColumn).value
	webbrowser.open(sheetValue)  # open the link in Internet Explorer
	# print(sheetValue)
	linkRow += 1  # increment by 1
print("All links have been opened.")  # finish