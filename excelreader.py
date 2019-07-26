# all for importing edited excel doc

# all for importing edited excel doc

import xlrd

file_loc = ("C:\Users\wblankenship\Documents\NIA Work Documents\FE\ASE SPT Protocol Translator Script\Book1.xlsx")
wb = xlrd.open_workbook(file_loc)
sheet = wb.sheet_by_name("Sheet1")
rows = sheet.nrows

b = [None] * rows
c = [None] * rows
d = [None] * rows
binarytags = [None] * rows
controltags = [None] * rows
analogtags = [None] * rows
controlTagsOperations = [None] * rows

rows2 = 0
rows3 = 0

# determing the number of filled rows to create row2, for control points. This is used for calculating how many inputs to add. This will have to be tested for different cases.
for x in range(sheet.nrows):
	cellValue = sheet.cell_value(x, 2)
	if cellValue != "":
		rows2 += 1

# determing the number of filled rows to create row3. This is used for calculating how many inputs to add. This will have to be tested for different cases.
for x in range(sheet.nrows):
	cellValue = sheet.cell_value(x, 3)
	if cellValue != "":
		rows3 += 1

# for determing the number of non empty cell values for the binary points column
for x in range(sheet.nrows):
	cellValue = sheet.cell_value(x, 1)
	if cellValue != "":
		b[x] = cellValue

# assigning values to c[0...max] for control points
# ********* there needs to be precautions put in if there is only one name instead of two
for x in range(rows2 / 2):
	cellValue = sheet.cell_value(x * 2, 2)
	if cellValue != "":
		c[x] = cellValue
	else:
		c[x] = None

# assigning values to d[0...max] for analog points
for x in range(rows3):
	cellValue = sheet.cell_value(x, 3)
	if cellValue != "":
		d[x] = cellValue
	else:
		d[x] = None

# collecting the control function from the book1 excel sheet
for x in range(rows2):
	cellValue = sheet.cell_value(x, 5)
	if cellValue != "":
		controlTagsOperations[x] = cellValue
	else:
		controlTagsOperations[x] = None

# number of groups to create for the control points
k = 0
for x in range(rows2):
	if (controlTagsOperations)

# loops for creating the client side control points based on their name and functions
for i in range(rows2):
	if controlTagsOperations[i] = "DISABLE":
		P5 = etree.Element("P", Id=str(i), Key=str(j), Name=str(binarytags[i]).replace(" ", "_").replace("\\", "_").replace("//", "_").replace("-", "_").upper())
		P5.text = ""
		P5.tail = "\n"
		group1.insert(i, P5)

tree = etree.ElementTree(root)
tree.write("sample2.xml")
# add for addtional info - - - > ,xml_declaration=True,   encoding="utf-8")






		

	

	



	

