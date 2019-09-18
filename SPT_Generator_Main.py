from lxml import etree
import xlrd
from Tkinter import *
from S9000 import s9000Gen
from TRW9550 import trw9550Gen

file_loc = ("C:\Users\wblankenship\Documents\GitHub\SPT-Builder\Working\Hornerstown S9000.xlsx")
wb = xlrd.open_workbook(file_loc)
sheet = wb.sheet_by_name("Sheet1")
rows = sheet.nrows

# ***************************************************************************************Taking user input with Tkinter
if __name__ == '__main__':
    root = Tk()

    variable = StringVar(root)
    variable.set("Yes") # default value
    variable2 = StringVar(root)
    variable2.set(1200)
    variable3 = StringVar(root)
    variable3.set(9600)
    variable4 = StringVar(root)
    variable4.set("TRW S9000")

    def my_function():

        global stationName
        global emsBaudRate
        global emsLocalAddress
        global rtuBaudRate
        global rtuRemoteAddress
        global rxTxUserInput
        global revision
        global protocol

        input = variable4.get()
        protocol = str(input)
        input1 = my_entry.get()
        stationName = str(input1)
        # [:5] limits input to only 5 character no matter how long the input is
        input2 = variable3.get()
        emsBaudRate = str(input2)
        input3 = my_entry2.get()[:5]
        emsLocalAddress = str(input3)
        input4 = variable2.get()
        rtuBaudRate = str(input4)
        input5 = my_entry4.get()[:5]
        rtuRemoteAddress = str(input5)
        input6 = variable.get()
        rxTxUserInput = str(input6.upper())
        input7 = my_entry6.get()[:1]
        revision = str(input7)
        root.destroy()


    # used for determining that user input is a digit
    def testVal(inStr,acttyp):
        if acttyp == '1': #insert
            if not inStr.isdigit() and inStr > 10000:
                return False
        return True

    def testVal2(inStr,acttyp):
        if acttyp == '1': #insert
            if inStr.isdigit():
                return False
        return True


    my_label = Label(root, text="Protocol")
    my_label.grid(row=0, column=0)
    my_entry7 = Entry(root, validate="key")
    # validates that user input is only digits with testVal
    my_entry7 = OptionMenu(root, variable4, "TRW S9000", "TRW 9550", "CDC 44-550")
    my_entry7.pack()
    my_entry7.grid(row=0, column=1)

    my_label = Label(root, text = "Station Name")
    my_label.grid(row = 1, column = 0)
    # validates that user input is only letters with testVal2
    my_entry = Entry(root, validate="key")
    my_entry['validatecommand'] = (my_entry.register(testVal2),'%P','%d')
    my_entry.pack()
    my_entry.grid(row = 1, column = 1)

    my_label = Label(root, text = "EMS Baud Rate")
    my_label.grid(row = 2, column = 0)
    my_entry1 = Entry(root, validate="key")
    # validates that user input is only digits with testVal
    my_entry1 = OptionMenu(root, variable3, "9600")
    my_entry1.pack()
    my_entry1.grid(row = 2, column = 1)

    my_label = Label(root, text = "To EMS Local Address (0-65519)")
    my_label.grid(row = 3, column = 0)
    my_entry2 = Entry(root, validate="key")
    my_entry2['validatecommand'] = (my_entry2.register(testVal),'%P','%d')
    my_entry2.pack()
    my_entry2.grid(row = 3, column = 1)

    my_label = Label(root, text = "RTU Baud Rate")
    my_label.grid(row = 4, column = 0)
    my_entry3 = OptionMenu(root, variable2, "1200", "2400")
    my_entry3.pack()
    my_entry3.grid(row = 4, column = 1)

    my_label = Label(root, text = "To RTU Remote Address (Example: 3)")
    my_label.grid(row = 5, column = 0)
    my_entry4 = Entry(root, validate="key")
    my_entry4['validatecommand'] = (my_entry4.register(testVal),'%P','%d')
    my_entry4.pack()
    my_entry4.grid(row = 5, column = 1)

    #changes whether the switched rx / tx = 0 for no modem, False or 1 for modem, True
    my_label = Label(root, text = "Switched RX and Switched TX - Is there a modem? Answer: 'Yes' or 'No'")
    my_label.grid(row = 6, column = 0)
    my_entry5 = OptionMenu(root, variable, "Yes", "No")
    my_entry5.pack()
    my_entry5.grid(row = 6, column = 1)

    my_label = Label(root, text = "Config Revision?")
    my_label.grid(row = 7, column = 0)
    my_entry6 = Entry(root, validate="key")
    my_entry6['validatecommand'] = (my_entry6.register(testVal2),'%P','%d')
    my_entry6.pack()
    my_entry6.grid(row = 7, column = 1)

    # Variables for data collected from the user to create the config
    stationName = ""
    emsBaudRate = ""
    emsLocalAddress = ""
    rtuBaudRate = ""
    rtuRemoteAddress = ""
    revision = ""
    rxTxUserInput = ""
    protocol = ""

    my_button = Button(root, text = "Submit", command = my_function)
    my_button.grid(row = 8, column = 1)
    root.title("SPT Generator")
    root.resizable(0, 0)
    root.mainloop()

    if protocol == "TRW S9000":
        s9000Gen(rxTxUserInput, stationName, revision, emsBaudRate, rtuBaudRate, emsLocalAddress, rtuRemoteAddress)
    if protocol == "TRW 9550":
        trw9550Gen(rxTxUserInput, stationName, revision, emsBaudRate, rtuBaudRate, emsLocalAddress, rtuRemoteAddress)
    if protocol == "CDC 44-550":
        print("")





# ***************************************************************************************Taking user input with Tkinter