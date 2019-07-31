#import tkinter as tk # Python 3 import
from Tkinter import * # Python 2 import

# root = tk.Tk()
#
# def my_function():
#     global myVar
#     input1 = my_entry.get()
#     myVar = input1
#     input2 = my_entry2.get()
#     input3 = my_entry3.get()
#     input4 = my_entry4.get()
#     input5 = my_entry5.get()
#     print(str(input1 + "\n" + input2 + "\n" + input3 + "\n" + input4 + "\n" + input5))
#     root.destroy()
#     #do stuff with url_member
#
#
#
# my_label = tk.Label(root, text = "EMS Baud Rate")
# my_label.grid(row = 0, column = 0)
# my_entry = tk.Entry(root)
# my_entry.grid(row = 0, column = 1)
# my_label = tk.Label(root, text = "To EMS Local Address (Example: 10237)")
# my_label.grid(row = 1, column = 0)
# my_entry2 = tk.Entry(root)
# my_entry2.grid(row = 1, column = 1)
# my_label = tk.Label(root, text = "RTU Baud Rate")
# my_label.grid(row = 2, column = 0)
# my_entry3 = tk.Entry(root)
# my_entry3.grid(row = 2, column = 1)
# my_label = tk.Label(root, text = "To RTU Remote Address (Example: 3)")
# my_label.grid(row = 3, column = 0)
# my_entry4 = tk.Entry(root)
# my_entry4.grid(row = 3, column = 1)
# #changes whether the switched rx / tx = 0 for no modem, False or 1 for modem, True
# my_label = tk.Label(root, text = "Switched RX and Switched TX - Is there a modem? Answer: 'Yes' or 'No'")
# my_label.grid(row = 4, column = 0)
# my_entry5 = tk.Entry(root)
# my_entry5.grid(row = 4, column = 1)
# myVar = ""
# my_button = tk.Button(root, text = "Submit", command = my_function)
# my_button.grid(row = 5, column = 1)
#
# root.mainloop()

master = Tk()

variable = StringVar(master)
variable.set("one") # default value

my_label = Label(master, text = "Station Name")
my_label.grid(row = 0, column = 0)\

w = OptionMenu(master, variable, "one", "two", "three")
w.pack()
w.grid(row = 0, column = 1)

def ok():
    userInput = variable.get()
    print ("value is:" + userInput)


button = Button(master, text="OK", command=ok)
button.pack()
button.grid(row = 1, column = 1)

mainloop()