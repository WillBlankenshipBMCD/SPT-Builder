# all for importing edited excel doc

#Extracting number of rows
import xlrd
file_loc = ("C:\Users\wblankenship\Documents\NIA Work Documents\FE\ASE SPT Protocol Translator Script\Book1.xlsx")

wb = xlrd.open_workbook(file_loc) 
sheet = wb.sheet_by_index(0) 
rows = sheet.nrows
b = [None] * 100
	
for i in range(sheet.nrows): 
	a = str(sheet.cell_value(i, 1))
	if a != "":
		b[i] = a
	print(b[i])
	raw_input()
		
		

	

	



	

