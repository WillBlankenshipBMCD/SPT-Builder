
from lxml import etree
import Tkinter as tk

#Extracting number of rows
import xlrd
file_loc = ("C:\Users\wblankenship\Documents\GitHub\SPT-Builder\Hornerstown S9000.xlsx")
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
# these lists are all used to keep track of the key values for all the points
binaryNums = [None] * rows
analogNums = [None] * rows
controlNums = [None] * rows

rows2 = 0
rows3 = 0

# ***************************************************************************************Taking user input with Tkinter
root = tk.Tk()

def my_function():
    global stationName
    global emsBaudRate
    global emsLocalAddress
    global rtuBaudRate
    global rtuRemoteAddress
    global rxTxUserInput

    input1 = my_entry.get()
    stationName = input1
    input2 = my_entry1.get()
    emsBaudRate = input2
    input3 = my_entry2.get()
    emsLocalAddress = input3
    input4 = my_entry3.get()
    rtuBaudRate = input4
    input5 = my_entry4.get()
    rtuRemoteAddress = input5
    input6 = my_entry5.get()
    rxTxUserInput = input6
    root.destroy()

my_label = tk.Label(root, text = "Station Name")
my_label.grid(row = 0, column = 0)
my_entry = tk.Entry(root)
my_entry.grid(row = 0, column = 1)

my_label = tk.Label(root, text = "EMS Baud Rate")
my_label.grid(row = 1, column = 0)
my_entry1 = tk.Entry(root)
my_entry1.grid(row = 1, column = 1)

my_label = tk.Label(root, text = "To EMS Local Address (Example: 10237)")
my_label.grid(row = 2, column = 0)
my_entry2 = tk.Entry(root)
my_entry2.grid(row = 2, column = 1)

my_label = tk.Label(root, text = "RTU Baud Rate")
my_label.grid(row = 3, column = 0)
my_entry3 = tk.Entry(root)
my_entry3.grid(row = 3, column = 1)

my_label = tk.Label(root, text = "To RTU Remote Address (Example: 3)")
my_label.grid(row = 4, column = 0)
my_entry4 = tk.Entry(root)
my_entry4.grid(row = 4, column = 1)

#changes whether the switched rx / tx = 0 for no modem, False or 1 for modem, True
my_label = tk.Label(root, text = "Switched RX and Switched TX - Is there a modem? Answer: 'Yes' or 'No'")
my_label.grid(row = 5, column = 0)
my_entry5 = tk.Entry(root)
my_entry5.grid(row = 5, column = 1)

# Variables for data collected from the user to create the config
stationName = ""
emsBaudRate = ""
emsLocalAddress = ""
rtuBaudRate = ""
rtuRemoteAddress = ""
rxTxUserInput = ""
RxTxAdded = ""

my_button = tk.Button(root, text = "Submit", command = my_function)
my_button.grid(row = 6, column = 1)

root.mainloop()


# ***************************************************************************************Taking user input with Tkinter

# determing the number of filled rows to create row2, for control points. This is used for calculating how many inputs to add. This will have to be tested for different cases. 
for x in range(sheet.nrows):
    cellValue = sheet.cell_value(x,2)
    if cellValue != "":
        rows2 += 1


# determing the number of filled rows to create row3. This is used for calculating how many analog inputs to add. This will have to be tested for different cases.
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
# emsBaudRate collected from user in Tkinter
baudrate.text = str(emsBaudRate)
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

#start of Device
# emsLocalAddress is collected from the user with Tkinter
device1 = etree.Element("DEVICE", Id = str(emsLocalAddress), Name = "toEMS")
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
# j is the variable that is used to increase the point ID for all the points on the Server and Client sides
j = 118
for i in range(rows):
    if b[i] != None:
        P = etree.Element("P", Id = str(i), Ref = str(j))
        P.text = ""
        P.tail = "\n"
        object1.insert(i, P)
        binaryNums[i] = j
    else:
        P = etree.Element("P", Id = str(i))
        P.text = ""
        P.tail = "\n"
        object1.insert(i, P)
        binaryNums[i] = j
    j += 1


#start of Object 2 - for Control Points ID = 12
object2 = etree.Element("OBJECT", Id = "12")
object2.text = "\n"
object2.tail = "\n"
device1.insert(1, object2)


i = 0
h = 0
k = 0
# range(rows2/2)+1 because it doubles the name of each and there is ususally two rows for each point
# this may have to be tweaked depending on whether the names are doubled or not
for i in range((rows2/2)):

    if c[i] == "UNDEFINED":

        P2 = etree.Element("P", Id=str(i))
        P2.text = "\n"
        P2.tail = "\n"
        object2.insert(i, P2)

        F1 = etree.Element("F", Id="0")
        F1.text = ""
        F1.tail = "\n"
        P2.insert(k, F1)
        k += 1
        j += 1
        controlNums[h] = j
        h += 1

        F2 = etree.Element("F", Id="1")
        F2.text = ""
        F2.tail = "\n"
        P2.insert(k, F2)
        k += 1
        j += 1
        controlNums[h] = j
        h += 1

    elif c[i] != None:

        P2 = etree.Element("P", Id=str(i))
        P2.text = "\n"
        P2.tail = "\n"
        object2.insert(i, P2)

        F1 = etree.Element("F", Id="0", Ref=str(j))
        F1.text = ""
        F1.tail = "\n"
        P2.insert(k, F1)
        k += 1
        controlNums[h] = j
        j += 1
        h += 1

        F2 = etree.Element("F", Id="1", Ref=str(j))
        F2.text = ""
        F2.tail = "\n"
        P2.insert(k, F2)
        k += 1
        controlNums[h] = j
        j += 1
        h += 1




# start of Object 3 - for Analog Points ID = 30
object3 = etree.Element("OBJECT", Id = "30")
object3.text = "\n"
object3.tail = "\n"
device1.insert(2, object3)

i = 0

# adding analog points on the server side
for i in range(rows3):
    if c[i] != None:
        P3 = etree.Element("P", Id = str(i), Ref = str(j))
        P3.text = ""
        P3.tail = "\n"
        object3.insert(i, P3)
        analogNums[i] = j
    else:
        P3 = etree.Element("P", Id = str(i))
        P3.text = ""
        P3.tail = "\n"
        object3.insert(i, P3)
        analogNums[i] = j

    j += 1

#used for all the group and object keys. These keys do not reference back to the server side. They can be any arbitrary value.
#I just started them off from where the point numbers that mattered ended.
objGroupNums = j + 1
j = 118

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
baudrate2.text = rtuBaudRate
baudrate2.tail = "\n"
line2.insert(2, baudrate2)

#where data needs to be collected from the user, i.e. Address for RTU, and Name
#start of Device
device2 = etree.Element("DEVICE", Id = rtuRemoteAddress, Key = "114", Name = stationName + "_TRWS9")
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

#used to start the next loop for the binary SS or NL points
l = 0
for i in range(rows):
    #keeping track of where to start for the SS points, this loop comes last for all the points
    l=i
    #checks if we have reached the end of the CD type binaries, this may need to be changed depending on order of points or point type
    if b[i] == "NL":
        break
    #checks if the cell is either empty or a NL binary point
    if b[i] != None and b[i] != "NL":
        P4 = etree.Element("P", Id = str(i), Key = str(binaryNums[i]), Name = str(binarytags[i]).strip().replace(" ", "_")
                           .replace("\\","_").replace("/","_").replace("-","_")
                           .replace("&", "").replace("<", "").replace(">", "")
                           .replace('"',"").upper())
        P4.text = ""
        P4.tail = "\n"
        group1.insert(i, P4)

    else:
        P4 = etree.Element("P", Id = str(i), Key = str(binaryNums[i]))
        P4.text = ""
        P4.tail = "\n"
        group1.insert(i, P4)



#start of Object 5 - Analog Inputs ID = 2 for all binary inputs MCD and SS or NL
object5 = etree.Element("OBJECT", Id = "2", Key = str(objGroupNums))
object5.text = "\n"
object5.tail = "\n"
device2.insert(2, object5)
objGroupNums+=1


# for putting all the analog points in a list
for x in range(rows3):
    cellValue = sheet.cell_value(x,3)
    if cellValue != "":
        analogtags[x] = cellValue
    else:
        analogtags[x] = cellValue

i = 0
k = 0
h = 0
j = 0
#loops for creating the client side analog points besed on their name and frame #
#creating each frame or "GROUP"
for i in range(rows3/4):

    analogGroups = etree.Element("GROUP", Id = str(i), Key = str(objGroupNums))
    analogGroups.text = "\n"
    analogGroups.tail = "\n"
    object5.insert(h,analogGroups)
    objGroupNums += 1

    for h in range (4):

        P5 = etree.Element("P", Id = str(h), Key = str(analogNums[j]),
                           Name = str(analogtags[k]).strip().replace(" ", "_")
                           .replace("\\","_").replace("/","_").replace("-","_")
                           .replace("&", "").replace("<", "").replace(">", "").replace('"',"").upper())
        P5.text = ""
        P5.tail = "\n"
        analogGroups.insert(h,P5)
        k+=1
        j+=1



#start of Object 6 - Control Inputs ID = 4 for all Control inputs MCD and SS or NL
object6 = etree.Element("OBJECT", Id = "4", Key = str(objGroupNums))
object6.text = "\n"
object6.tail = "\n"
device2.insert(3, object6)
objGroupNums += 1

a = rows2 // 16
e = rows2 % 16
controlNumGroups = a
if e != 0:
    controlNumGroups += 1

# for putting all the control points in a list
for x in range(rows2):
    cellValue = sheet.cell_value(x,2)
    if cellValue != "":
        controltags[x] = cellValue
    else:
        controltags[x] = cellValue

a = 0
h = 0
j = 0
k = 0
for i in range(controlNumGroups):
    controlGroups = etree.Element("GROUP", Id=str(i+1), Key=str(objGroupNums))
    controlGroups.text = "\n"
    controlGroups.tail = "\n"
    object6.insert(h, controlGroups)
    objGroupNums += 1

    for h in range(16):

        if controltags[k] == "UNDEFINED":
            P6 = etree.Element("P", Id=str(h), Key=str(controlNums[j]))
            P6.text = ""
            P6.tail = "\n"
            controlGroups.insert(h, P6)
        elif controltags[k].upper() == "SPARE":
            P6 = etree.Element("P", Id=str(h), Key=str(controlNums[j]), Name=str(controltags[k]).strip().replace(" ", "_")
                               .replace("\\", "_").replace("/","_").replace("-", "_")
                               .replace("&", "").replace("<", "").replace(">", "").replace('"',"").upper())
            P6.text = ""
            P6.tail = "\n"
            controlGroups.insert(h, P6)
        elif controltags[k] != "":
            P6 = etree.Element("P", Id=str(h), Key=str(controlNums[j]), Name=str(controltags[k]).strip().replace(" ", "_").replace("\\", "_")
                               .replace("/", "_").replace("-", "_").replace("&", "")
                               .replace("<", "").replace(">", "").replace('"',"").upper() + "_" +
                                str(sheet.cell_value(a,5).replace(" ", "_").replace("&", "")
                                .replace("<", "").replace(">", "").replace('"',"")))
            P6.text = ""
            P6.tail = "\n"
            controlGroups.insert(h, P6)
        if sheet.cell_value(a, 5).upper() == "DISABLE":
            invert = etree.Element("Invert")
            invert.text = "0"
            invert.tail = "\n"
            P6.insert(0,invert)
            P6.text = "\n"
        a += 1
        k += 1
        j += 1




#start of Object 7 - Binary Inputs ID = 4 for all Control inputs MCD and SS or NL
object7 = etree.Element("OBJECT", Id = "7", Key = str(objGroupNums))
object7.text = "\n"
object7.tail = "\n"
device2.insert(4, object7)
objGroupNums += 1

#start of Group 7 - Binary Inputs ID = 0 for all Control inputs MCD and SS or NL
group7 = etree.Element("GROUP", Id = "0", Key = str(objGroupNums))
group7.text = "\n"
group7.tail = "\n"
object7.insert(0, group7)
objGroupNums += 1

for i in range(rows):

    #checks if we have reached the end of the CD type binaries
    if b[i] == "internal":
        break
    #checks if the cell is either empty or a NL binary point
    if b[i] == None:
        P7 = etree.Element("P", Id=str(l), Key=str(binaryNums[l]))
        P7.text = ""
        P7.tail = "\n"
        group7.insert(i, P7)
        l += 1

    elif b[i] != "CD":
        P7 = etree.Element("P", Id=str(l), Key=str(binaryNums[l]), Name=str(binarytags[l]).strip().replace(" ", "_").replace("\\", "_")
                           .replace("/", "_").replace("-","_").replace("  ", "_")
                           .replace("&", "").replace("<", "").replace(">", "").replace('"',"").upper())
        P7.text = ""
        P7.tail = "\n"
        P7.tail = "\n"
        group7.insert(i, P7)
        l += 1

#start of Object 8 - Binary Inputs ID = 4 for all Control inputs MCD and SS or NL
object8 = etree.Element("OBJECT", Id = "16384", Key = str(objGroupNums))
object8.text = "\n"
object8.tail = "\n"
device2.insert(5, object8)

P8 = etree.Element("P", Id = "0", Key=str(binaryNums[l]), Name=str(binarytags[l]).strip().replace(" ", "_")
                   .replace("\\", "_").replace("/", "_").replace("-","_")
                   .replace("  ", "_").replace("&", "").replace("<", "")
                   .replace(">", "").replace('"',"").upper())
P8.text = ""
P8.tail = "\n"
object8.insert(0, P8)


tree = etree.ElementTree(root)
tree.write(stationName + " ASE SPT Alert 9000 DNP Rev " + "A" + ".xml")
#add for addtional info - - - > ,xml_declaration=True,   encoding="utf-8")  

