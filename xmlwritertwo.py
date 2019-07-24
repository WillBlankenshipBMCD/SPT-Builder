
from lxml import etree

#Extracting number of rows
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
	cellValue = sheet.cell_value(x,2)
	if cellValue != "":
		rows2 += 1

# determing the number of filled rows to create row3. This is used for calculating how many inputs to add. This will have to be tested for different cases. 
for x in range(sheet.nrows):
	cellValue = sheet.cell_value(x,3)
	if cellValue != "":
		rows3 += 1

# for determing the number of non empty cell values for the binary points column 
for x in range(sheet.nrows):
	cellValue = sheet.cell_value(x,1)
	if cellValue != "":
		b[x] = cellValue

# assigning values to c[0...max] for control points
# ********* there needs to be precautions put in if there is only one name instead of two 
for x in range(rows2/2):
	cellValue = sheet.cell_value(x*2,2)
	if cellValue != "":
		c[x] = cellValue
	else:
		c[x] = None

# assigning values to d[0...max] for analog points
for x in range(rows3):
	cellValue = sheet.cell_value(x,3)
	if cellValue != "":
		d[x] = cellValue
	else:
		d[x] = None

# collecting the control function from the book1 excel sheet
for x in range(rows2):
	cellValue = sheet.cell_value(x,5)
	if cellValue != "":
		controlTagsOperations[x] = cellValue
	else:
		controlTagsOperations[x] = None
		


#root of the document
root = etree.Element("SPT", Id = "1")
root.text = "\n"


#root.set('Ref','18')

#start of direction1 for Server side portion
direction = etree.Element("DIRECTION", Id = "0")
direction.text = "\n"
direction.tail = "\n"
root.insert(0,direction)

#start of protocol1 for server side portion
protocol = etree.Element("PROTOCOL", Id = "1")
protocol.text = "\n"
protocol.tail = "\n"
direction.insert(0, protocol)

#start of line1 for server side portion
line = etree.Element("LINE", Id = "1" )
line.text = "\n"
line.tail = "\n"
protocol.insert(0, line)

#start of interface for server side portion
interface = etree.Element("Interface")
interface.text = "0"
interface.tail = "\n"
line.insert(1, interface)

#start of baudrate for server side portion
baudrate = etree.Element("BaudRate")
baudrate.text = "9600"
baudrate.tail = "\n"
line.insert(2, baudrate)

#start of SwitchedRx just for server side portion
switchedRx = etree.Element("SwitchedRX")
switchedRx.text = "0"
switchedRx.tail = "\n"
line.insert(3, switchedRx)

#start of SwitchedTx just for server side portion
switchedTx = etree.Element("SwitchedTX")
switchedTx.text = "0"
switchedTx.tail = "\n"
line.insert(4, switchedTx)

#where data needs to start being collected from the user, i.e. Address for EMS - it's always toEMS
#start of Device
device1 = etree.Element("DEVICE", Id = "10237", Name = "toEMS")
device1.text = "\n"
device1.tail = "\n"
line.insert(5, device1)

#start of Object 1 - Binary Inputs first ID = 1 for all binary inputs MCD and SS or NL
object1 = etree.Element("OBJECT", Id = "1")
object1.text = "\n"
object1.tail = "\n"
device1.insert(0, object1)

# Where data needs to be inserted for binary inputs 
# a loop to increase the ID can be used here
# a loop to match the ref/key pair can be created. 111 through how ever many tags are added on to the RTU / Client side
# for s9000 first tag from what I currently know is 118, could change if settings are changed. I need to verify this.
# possibly need to create the client side first before the server side toEMS in order to get the proper numbering of points based on the names and the 
# for 9550
# for CDC
# start of P0 or Binary 0

# loop for generating the binary tags on the server side
i = 0
j = 118
for i in range(rows):
	if b[i] != None:
		P = etree.Element("P", Id = str(i), Ref = str(j))
		P.text = ""
		P.tail = "\n"
		object1.insert(i, P)
	else: 
		P = etree.Element("P", Id = str(i))
		P.text = ""
		P.tail = "\n"
		object1.insert(i, P)
	j += 1


#start of Object 2 - for Control Points ID = 12
object2 = etree.Element("OBJECT", Id = "12")
object2.text = "\n"
object2.tail = "\n"
device1.insert(1, object2)


i = 0
j = 150
k = 0
# range(rows2/2)+1 because it doubles the name of each and there is ususally two rows for each point
# this may have to be tweaked depending on whether the names are doubled or not
for i in range((rows2/2)):
	
	if c[i] != None:
		
		P2 = etree.Element("P", Id = str(i))
		P2.text = "\n"
		P2.tail = "\n"
		object2.insert(i, P2)
		
		F1 = etree.Element("F", Id = "0", Ref = str(j))
		F1.text = ""
		F1.tail = "\n"
		P2.insert(k, F1)
		j+=1
		k+=1
		
		F2 = etree.Element("F", Id = "1", Ref = str(j))
		F2.text = ""
		F2.tail = "\n"
		P2.insert(k, F2)
		j+=1
		k+=1	
	else: 
		
		P2 = etree.Element("P", Id = str(i))
		P2.text = "\n"
		P2.tail = "\n"
		object2.insert(i, P2)
		
		F1 = etree.Element("F", Id = "0")
		F1.text = ""
		F1.tail = "\n"
		P2.insert(k, F1)
		k+=1
		
		F2 = etree.Element("F", Id = "1")
		F2.text = ""
		F2.tail = "\n"
		P2.insert(k, F2)
		k+=1
		
	
# start of Object 3 - for Analog Points ID = 30
object3 = etree.Element("OBJECT", Id = "30")
object3.text = "\n"
object3.tail = "\n"
device1.insert(2, object3)

i = 0
j = 134
# points come in 4's. The other points get added in the client application side 
for i in range(rows3):
	if c[i] != None:
		if i != 0 and (i%4) == 0:
			j+=1
		P3 = etree.Element("P", Id = str(i), Ref = str(j))
		P3.text = ""
		P3.tail = "\n"
		object3.insert(i, P3)
		
	else: 
		if i != 0 and (i%4) == 0:
			j+=1
		P3 = etree.Element("P", Id = str(i))
		P3.text = ""
		P3.tail = "\n"
		object3.insert(i, P3)
		
	j += 1

#start of direction2 for client side portion
direction2 = etree.Element("DIRECTION", Id = "1", Key = "111")
direction2.text = "\n"
direction2.tail = "\n"
root.insert(1,direction2)

#start of protocol2 for server side portion
protocol2 = etree.Element("PROTOCOL", Id = "9", Key = "112")
protocol2.text = "\n"
protocol2.tail = "\n"
direction2.insert(0, protocol2)

#start of line1 for server side portion
line2 = etree.Element("LINE", Id = "2", Key = "113" )
line2.text = "\n"
line2.tail = "\n"
protocol2.insert(0, line2)

#start of interface for server side portion
interface2 = etree.Element("Interface")
interface2.text = "0"
interface2.tail = "\n"
line2.insert(1, interface2)

#start of baudrate for server side portion
baudrate2 = etree.Element("BaudRate")
baudrate2.text = "1200"
baudrate2.tail = "\n"
line2.insert(2, baudrate2)

#where data needs to be collected from the user, i.e. Address for RTU, and Name
#start of Device
device2 = etree.Element("DEVICE", Id = "3", Key = "114", Name = "Millhurst_TRWS9")
device2.text = "\n"
device2.tail = "\n"
line2.insert(3, device2)

#start of request 
request = etree.Element("REQUEST", Id = "0", Key = "115")
request.text = "\n"
request.tail = "\n"
device2.insert(0, request)

# the next four have to do specifically with settings for S9000
#start of function
function = etree.Element("Function")
function.text = "5"
function.tail = "\n"
request.insert(0, function)

#start of variation
variation = etree.Element("Variation")
variation.text = "1"
variation.tail = "\n"
request.insert(1, variation)

#start of XCount
xcount = etree.Element("XCount")
xcount.text = "1"
xcount.tail = "\n"
request.insert(2, xcount)

#start of YCount
ycount = etree.Element("YCount")
ycount.text = "3"
ycount.tail = "\n"
request.insert(3, ycount)

#start of Object 4 - Binary Inputs first ID = 1 for all binary inputs MCD and SS or NL
object4 = etree.Element("OBJECT", Id = "1", Key = "116")
object4.text = "\n"
object4.tail = "\n"
device2.insert(1, object4)

#start of Object 4 - Binary Inputs first ID = 1 for MCD on the Client side
group1 = etree.Element("GROUP", Id = "0", Key = "117")
group1.text = "\n"
group1.tail = "\n"
object4.insert(0, group1)

# for putting all the binary points in a list
for x in range(sheet.nrows):
	cellValue = sheet.cell_value(x,0)
	if cellValue != "":
		binarytags[x] = cellValue
	else:
		binarytags[x] = cellValue

# loop for generating the binary mcd/cd tags on the client side
i = 0
j = 118
#used to start the next loop for the binary SS or NL points
k = 0
for i in range(rows):
	#keeping track of where to start for the SS points, this loop comes last for all the points
	k=i
	#checks if we have reached the end of the CD type binaries
	if b[i] == "NL":
		break
	#checks if the cell is either empty or a NL binary point
	if b[i] != None and b[i] != "NL":
		P4 = etree.Element("P", Id = str(i), Key = str(j), Name = str(binarytags[i]).replace(" ", "_").replace("\\","_").replace("//","_").replace("-","_").upper())
		P4.text = ""
		P4.tail = "\n"
		group1.insert(i, P4)
	else: 
		P4 = etree.Element("P", Id = str(i), Key = str(j))
		P4.text = ""
		P4.tail = "\n"
		group1.insert(i, P4)
	j += 1
	
#start of Object 5 - Analog Inputs ID = 2 for all binary inputs MCD and SS or NL
object5 = etree.Element("OBJECT", Id = "2", Key = "132")
object5.text = "\n"
object5.tail = "\n"
device2.insert(2, object5)

i = 0
k = 0
h = 0
j = 133
#loops for creating the client side analog points besed on their name and frame #
#creating each frame or "GROUP"
for i in range(rows3/4):

	analogGroups = etree.Element("GROUP", Id = str(i), Key = str(j))
	analogGroups.text = "\n"
	analogGroups.tail = "\n"
	object5.insert(h,analogGroups)
	
	for h in range (4):
		j+=1
		P5 = etree.Element("P", Id = str(h), Key = str(j), Name = str(d[k]).replace(" ", "_").replace("\\","_").replace("//","_").replace("-","_").upper())
		P5.text = ""
		P5.tail = "\n"
		analogGroups.insert(h,P5)
		k+=1
		
	j+=1

#start of Object 6 - Control Inputs ID = 4 for all Control inputs MCD and SS or NL
object6 = etree.Element("OBJECT", Id = "4", Key = "148")
object6.text = "\n"
object6.tail = "\n"
device2.insert(3, object6)




			
tree = etree.ElementTree(root)
tree.write("sample.xml")
#add for addtional info - - - > ,xml_declaration=True,   encoding="utf-8")  

