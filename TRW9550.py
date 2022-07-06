from lxml import etree
from Tkinter import *
import pprint
pp = pprint.PrettyPrinter(indent=4)
import xlrd
file_loc = ("C:\Users\wblankenship\Documents\GitHub\SPT-Builder\SPT Config Generator.xlsx")
wb = xlrd.open_workbook(file_loc)
sheet = wb.sheet_by_name("Sheet1")
rows = sheet.nrows

def trw9550Gen(rxTxUserInput, stationName, revision, emsBaudRate, rtuBaudRate, emsLocalAddress, rtuRemoteAddress, EnableUmode, port):

    mcdPoints = [None] * rows
    ssPoints = [None] * rows
    controlPoints = [None] * rows
    analogPoints = [None] * rows

    # these lists are all used to keep track of the key values for all the points
    mcdBinaryNums = [None] * rows
    ssBinaryNums = [None] * rows
    analogNums = [None] * rows
    controlNums = [None] * rows

    # sorting and filling in the MCD Binary points
    for i in range(3, sheet.nrows):
        point = {"frame": sheet.cell_value(i, 1), "index": str(sheet.cell_value(i, 2)).replace(".0", ""),
                 "description": str(sheet.cell_value(i, 3)).strip().replace(" ", "_")
                     .replace("\\", "_").replace("/", "_").replace("-", "_")
                     .replace("&", "").replace("<", "").replace(">", "")
                     .replace('"', "").replace('._', "_").replace('_.', "_").replace('__', "_").replace('___',
                                                                                                        "_").upper()}

        if (point.get('index') == "" or point.get('frame') == ""):
            break

        mcdPoints[i - 3] = point

    # removing none values from the mcdPoints list
    mcdPoints = [i for i in mcdPoints if i]
    mcdPoints = sorted(mcdPoints, key=lambda i: (int(i['frame']), int(i['index'])))
    mcdPoints = [i for i in mcdPoints if i['description'] != ""]

    # sorting and filling in the SS Binary points
    for i in range(3, sheet.nrows):
        point = {"frame": str(sheet.cell_value(i, 5)).replace(".0", ""),
                 "index": str(sheet.cell_value(i, 6)).replace(".0", ""),
                 "description": str(sheet.cell_value(i, 7)).strip().replace(' "', "").replace('" ', "").replace(" ","_")
                     .replace("\\", "_").replace("/", "_").replace("-", "_")
                     .replace("&", "").replace("<", "").replace(">", "")
                     .replace('._', "_").replace('_.', "_").replace('__', "_").replace('___', "_").upper()}
        if (point.get('index') == "" or point.get('frame') == ""):
            break

        ssPoints[i - 3] = point
        # removing none values from the mcdPoints list
    ssPoints = [i for i in ssPoints if i]
    ssPoints = sorted(ssPoints, key=lambda i: (int(i['frame']), int(i['index'])))
    ssPoints = [i for i in ssPoints if i['description'] != ""]

    # sorting and filling in the Control points
    for i in range(3, sheet.nrows):
        if sheet.cell_value(i, 11) == "":
            index = str(sheet.cell_value(i - 1, 10)).replace(".0", "")
        else:
            index = str(sheet.cell_value(i, 10)).replace(".0", "")
        point = {"group": sheet.cell_value(i, 10), "index": index,
                 "description": str(sheet.cell_value(i, 11)).strip().replace(" ", "_")
                     .replace("\\", "_").replace("/", "_").replace("-", "_")
                     .replace("&", "").replace("<", "").replace(">", "")
                     .replace('"', "").replace('._', "_").replace('_.', "_").replace('__', "_").replace('___',"_").upper(),
                 "function": sheet.cell_value(i, 12)}
        if (point.get('group') == "" or point.get('index') == ""):
            break
        controlPoints[i - 3] = point

        # removing none values from the controlPoints list
    controlPoints = [i for i in controlPoints if i]
    controlPoints = sorted(controlPoints, key=lambda i: (int(i['group']), int(i['index'])))
    numDelete = 0

    for i in range(controlPoints.__len__()):
        if controlPoints[i].get('description') == "":
            numDelete += 1

    # if numDelete = 0, it will delete all the points in the list. It's just what [:0] does for some reason. This was a special case
    if numDelete != 0:
        controlPoints = controlPoints[:-numDelete]

    # filling in the Analog points
    for i in range(3, sheet.nrows):
        point = {"frame": str(sheet.cell_value(i, 14)).replace(".0", ""),
                 "index": str(sheet.cell_value(i, 15)).replace(".0", ""),
                 "description": str(sheet.cell_value(i, 16)).strip().replace(" ", "_")
                     .replace("\\", "_").replace("/", "_").replace("-", "_")
                     .replace("&", "").replace("<", "").replace(">", "")
                     .replace('"', "").replace('._', "_").replace('_.', "_").replace('__', "_").replace('___',
                                                                                                        "_").upper()}

        analogPoints[i - 3] = point
        # removing none values from the mcdPoints list

    analogPoints = [i for i in analogPoints if i]
    analogPoints = sorted(analogPoints, key=lambda i: (['frame'], ['index']))
    numDelete = 0
    for i in range(analogPoints.__len__()):
        if analogPoints[i].get('description') == "":
            numDelete += 1

    # if numDelete = 0, it will delete all the points in the list. It's just what [:0] does for some reason. This was a special case
    if numDelete != 0:
        analogPoints = analogPoints[:-numDelete]

    # root of the document
    root = etree.Element("SPT", Id="1")
    root.text = "\n"

    # root.set('Ref','18')

    # start of direction1 for Server side portion
    direction = etree.Element("DIRECTION", Id="0")
    direction.text = "\n"
    direction.tail = "\n"
    root.insert(0, direction)

    # start of protocol1 for server side portion
    protocol = etree.Element("PROTOCOL", Id="1")
    protocol.text = "\n"
    protocol.tail = "\n"
    direction.insert(0, protocol)

    # For enable u mode.
    if EnableUmode == "Yes":
        # start of Primary Adresss for Umode. This is the address to EMS.
        PrimaryAddress = etree.Element("PrimaryAddress")
        PrimaryAddress.text = "10"
        PrimaryAddress.tail = "\n"
        protocol.insert(0, PrimaryAddress)

        # start of Startup for Umode. This is the actual setting that enables U mode
        Startup = etree.Element("Startup")
        Startup.text = "1"
        Startup.tail = "\n"
        protocol.insert(1, Startup)

    # start of line1 for server side portion
    line = etree.Element("LINE", Id="1")
    line.text = "\n"
    line.tail = "\n"
    protocol.insert(2, line)

    # start of interface for server side portion
    interface = etree.Element("Interface")
    interface.text = "0"
    interface.tail = "\n"
    line.insert(1, interface)

    # start of baudrate for server side portion
    baudrate = etree.Element("BaudRate")
    # 9600 collected from user in Tkinter
    baudrate.text = str(emsBaudRate)
    baudrate.tail = "\n"
    line.insert(2, baudrate)

    # start of SwitchedRx just for server side portion
    switchedRx = etree.Element("SwitchedRX")
    switchedRx.text = "0"
    switchedRx.tail = "\n"
    line.insert(3, switchedRx)

    # start of SwitchedTx just for server side portion
    switchedTx = etree.Element("SwitchedTX")
    switchedTx.text = "0"
    switchedTx.tail = "\n"
    line.insert(4, switchedTx)

    # start of Device
    # 10235 is collected from the user with Tkinter
    device1 = etree.Element("DEVICE", Id=str(emsLocalAddress), Name="toEMS")
    device1.text = "\n"
    device1.tail = "\n"
    line.insert(5, device1)

    # start of Object 1 - Binary Inputs first ID = 1 for all binary inputs MCD and SS or NL
    object1 = etree.Element("OBJECT", Id="1")
    object1.text = "\n"
    object1.tail = "\n"
    device1.insert(0, object1)


    # j is the variable that is used to increase the point ID for all the points on the Server and Client sides
    j = 188
    id = 0

    if str(sheet.cell_value(3, 2)).replace(".0", "") == "0":

        for i in range(mcdPoints.__len__()):
            if mcdPoints[i].get('description') != "UNDEFINED":
                P = etree.Element("P", Id=str(id), Ref=str(j))
                P.text = ""
                P.tail = "\n"
                object1.insert(j, P)
                mcdBinaryNums[i] = j
            else:
                P = etree.Element("P", Id=str(id))
                P.text = ""
                P.tail = "\n"
                object1.insert(j, P)
                mcdBinaryNums[i] = j
            j += 1
            id += 1

        for i in range(ssPoints.__len__()):
            if ssPoints[i].get('description').upper() != "UNDEFINED":
                P = etree.Element("P", Id=str(id), Ref=str(j))
                P.text = ""
                P.tail = "\n"
                object1.insert(j, P)
                ssBinaryNums[i] = j
            else:
                P = etree.Element("P", Id=str(id))
                P.text = ""
                P.tail = "\n"
                object1.insert(j, P)
                ssBinaryNums[i] = j
            j += 1
            id += 1
    else:

        for i in range(ssPoints.__len__()):
            if ssPoints[i].get('description').upper() != "UNDEFINED":
                P = etree.Element("P", Id=str(id), Ref=str(j))
                P.text = ""
                P.tail = "\n"
                object1.insert(j, P)
                ssBinaryNums[i] = j
            else:
                P = etree.Element("P", Id=str(id))
                P.text = ""
                P.tail = "\n"
                object1.insert(j, P)
                ssBinaryNums[i] = j
            j += 1
            id += 1

        for i in range(mcdPoints.__len__()):
            if mcdPoints[i].get('description') != "UNDEFINED":
                P = etree.Element("P", Id=str(id), Ref=str(j))
                P.text = ""
                P.tail = "\n"
                object1.insert(j, P)
                mcdBinaryNums[i] = j
            else:
                P = etree.Element("P", Id=str(id))
                P.text = ""
                P.tail = "\n"
                object1.insert(j, P)
                mcdBinaryNums[i] = j
            j += 1
            id += 1

    # start of Object 2 - for Control Points ID = 12
    object2 = etree.Element("OBJECT", Id="12")
    object2.text = "\n"
    object2.tail = "\n"
    device1.insert(1, object2)

    id = 0
    h = 0
    id = 0
    h = 0
    for i in range(controlPoints.__len__() / 2):

        # i * 2 to account for there being 2 of each point
        if controlPoints[i * 2].get('description').upper() == "UNDEFINED":

            P = etree.Element("P", Id=str(id))
            P.text = "\n"
            P.tail = "\n"
            object2.insert(id, P)

            F1 = etree.Element("F", Id="0")
            F1.text = ""
            F1.tail = "\n"
            P.insert(j, F1)
            controlNums[h] = j
            h += 1
            j += 1

            F2 = etree.Element("F", Id="1")
            F2.text = ""
            F2.tail = "\n"
            P.insert(j, F2)
            controlNums[h] = j
            h += 1
            j += 1

        else:

            P = etree.Element("P", Id=str(id))
            P.text = "\n"
            P.tail = "\n"
            object2.insert(id, P)

            F1 = etree.Element("F", Id="0", Ref=str(j))
            F1.text = ""
            F1.tail = "\n"
            P.insert(j, F1)
            controlNums[h] = j
            h += 1
            j += 1

            F2 = etree.Element("F", Id="1", Ref=str(j))
            F2.text = ""
            F2.tail = "\n"
            P.insert(j, F2)
            controlNums[h] = j
            h += 1
            j += 1

        id += 1

    # start of Object 2 - for Control Points ID = 12
    object3 = etree.Element("OBJECT", Id="30")
    object3.text = "\n"
    object3.tail = "\n"
    device1.insert(2, object3)

    for i in range(analogPoints.__len__()):

        if analogPoints[i].get("description") != "":

            P3 = etree.Element("P", Id=str(i), Ref=str(j))
            P3.text = ""
            P3.tail = "\n"
            object3.insert(i, P3)
            analogNums[i] = j
        else:
            P3 = etree.Element("P", Id=str(i))
            P3.text = ""
            P3.tail = "\n"
            object3.insert(i, P3)
            analogNums[i] = j

        j += 1

    # start of direction2 for client side portion
    direction2 = etree.Element("DIRECTION", Id="1", Key="106")
    direction2.text = "\n"
    direction2.tail = "\n"
    root.insert(1, direction2)

    # start of protocol2 for server side portion. This Tag's ID determines which protocol it is.
    protocol2 = etree.Element("PROTOCOL", Id="49", Key="107")
    protocol2.text = "\n"
    protocol2.tail = "\n"
    direction2.insert(0, protocol2)

    # start of line2 for server side portion. For ID=port, port is directly reflective of the the port the RTU is connected to.
    # The name "line2" has no reference to the port. It's just the name of a variable.
    line2 = etree.Element("LINE", Id=port, Key="108")
    line2.text = "\n"
    line2.tail = "\n"
    protocol2.insert(0, line2)

    # start of interface for server side portion
    interface2 = etree.Element("Interface")
    interface2.text = "0"
    interface2.tail = "\n"
    line2.insert(1, interface2)

    # start of baudrate for server side portion
    baudrate2 = etree.Element("BaudRate")
    baudrate2.text = str(rtuBaudRate)
    baudrate2.tail = "\n"
    line2.insert(2, baudrate2)

    # added to every config, pre transmission delay is not added because default is already 20
    PostTransDelay = etree.Element("PostTransmissionDelay")
    PostTransDelay.text = "1"
    PostTransDelay.tail = "\n"
    line2.insert(3, PostTransDelay)

    # start of SwitchedRX and SwitchedTX  for server side portion
    if rxTxUserInput == "YES":
        SwitchedRX = etree.Element("SwitchedRX")
        SwitchedRX.text = "1"
        SwitchedRX.tail = "\n"
        line2.insert(4, SwitchedRX)
        SwitchedTX = etree.Element("SwitchedTX")
        SwitchedTX.text = "1"
        SwitchedTX.tail = "\n"
        line2.insert(5, SwitchedTX)
    else:
        SwitchedRX = etree.Element("SwitchedRX")
        SwitchedRX.text = "0"
        SwitchedRX.tail = "\n"
        line2.insert(4, SwitchedRX)
        SwitchedTX = etree.Element("SwitchedTX")
        SwitchedTX.text = "0"
        SwitchedTX.tail = "\n"
        line2.insert(5, SwitchedTX)

    # where data needs to be collected from the user, i.e. Address for RTU, and Name
    # start of Device
    device2 = etree.Element("DEVICE", Id=str(rtuRemoteAddress), Key="109", Name=str(stationName).replace(" ", "_") + "_TRW_9550")
    device2.text = "\n"
    device2.tail = "\n"
    line2.insert(6, device2)

    # start of request
    request = etree.Element("REQUEST", Id="0", Key="110")
    request.text = "\n"
    request.tail = "\n"
    device2.insert(0, request)


    # start of function
    function = etree.Element("Function")
    function.text = "5"
    function.tail = "\n"
    request.insert(0, function)

    # start of variation
    variation = etree.Element("Variation")
    variation.text = "1"
    variation.tail = "\n"
    request.insert(1, variation)

    # start of request
    request2 = etree.Element("REQUEST", Id="1", Key="111")
    request2.text = "\n"
    request2.tail = "\n"
    device2.insert(1, request2)

    # the next four have to do specifically with settings for TRW9550
    # start of function
    function = etree.Element("Function")
    function.text = "5"
    function.tail = "\n"
    request2.insert(0, function)

    # start of variation
    variation1 = etree.Element("Frequency")
    variation1.text = "5000"
    variation1.tail = "\n"
    request2.insert(1, variation1)

    # start of count
    count = etree.Element("Count")
    count.text = str(analogPoints.__len__())
    count.tail = "\n"
    request2.insert(2, count)

    # start of variation
    variation2 = etree.Element("Variation")
    variation2.text = "2"
    variation2.tail = "\n"
    request2.insert(3, variation2)

    # start of Object 4 - Binary Inputs first ID = 1 for all binary inputs MCD and SS or NL
    object4 = etree.Element("OBJECT", Id="1", Key="112")
    object4.text = "\n"
    object4.tail = "\n"
    device2.insert(2, object4)

    mcdID = 0

    if mcdPoints.__len__() > 28:

        # start of Object 4 - Binary Inputs first ID = 1 for MCD on the Client side
        group1 = etree.Element("GROUP", Id="0", Key="113")
        group1.text = "\n"
        group1.tail = "\n"
        object4.insert(0, group1)

        for i in range(28):
            if mcdPoints[i].get("description") == "SPT_COMM_FAIL":
                print("")
                mcdID += 1
            if mcdPoints[i].get('description').upper() != "UNDEFINED":
                P = etree.Element("P", Id=str(mcdID), Key=str(mcdBinaryNums[i]),
                                  Name=mcdPoints[i].get('description').strip().replace(" ", "_")
                                  .replace("\\", "_").replace("/", "_").replace("-", "_")
                                  .replace("&", "").replace("<", "").replace(">", "")
                                  .replace('"', "").upper())
                P.text = ""
                P.tail = "\n"
                group1.insert(i, P)
                mcdID += 1
            else:
                P = etree.Element("P", Id=str(i))
                P.text = ""
                P.tail = "\n"
                group1.insert(i, P)
                mcdID += 1

        if mcdPoints[28].get("description") == "SPT_COMM_FAIL":
            print("")
        else:
            group2 = etree.Element("GROUP", Id="1", Key="1200")
            group2.text = "\n"
            group2.tail = "\n"
            object4.insert(1, group2)
            for i in range(28, mcdPoints.__len__()):
                if mcdPoints[i].get("description") == "SPT_COMM_FAIL":
                    print("")
                    mcdID += 1
                elif mcdPoints[i].get('description').upper() != "UNDEFINED":
                    P = etree.Element("P", Id=str(mcdID), Key=str(mcdBinaryNums[i]),
                                      Name=mcdPoints[i].get('description').strip().replace(" ", "_")
                                      .replace("\\", "_").replace("/", "_").replace("-", "_")
                                      .replace("&", "").replace("<", "").replace(">", "")
                                      .replace('"', "").upper())
                    P.text = ""
                    P.tail = "\n"
                    group2.insert(i, P)
                    mcdID += 1
                else:
                    P = etree.Element("P", Id=str(i))
                    P.text = ""
                    P.tail = "\n"
                    object2.insert(i, P)
                mcdID += 1

    else:

        # start of Object 4 - Binary Inputs first ID = 1 for MCD on the Client side
        group1 = etree.Element("GROUP", Id="0", Key="117")
        group1.text = "\n"
        group1.tail = "\n"
        object4.insert(0, group1)

        for i in range(mcdPoints.__len__()):
            if mcdPoints[i].get("description") == "SPT_COMM_FAIL":
                print("")
                mcdID += 1
            elif mcdPoints[i].get('description').upper() != "UNDEFINED":

                P = etree.Element("P", Id=str(mcdID), Key=str(mcdBinaryNums[i]),
                                  Name=mcdPoints[i].get('description').strip().replace(" ", "_")
                                  .replace("\\", "_").replace("/", "_").replace("-", "_")
                                  .replace("&", "").replace("<", "").replace(">", "")
                                  .replace('"', "").upper())
                P.text = ""
                P.tail = "\n"
                group1.insert(i, P)
                mcdID += 1
            else:
                P = etree.Element("P", Id=str(mcdID))
                P.text = ""
                P.tail = "\n"
                group1.insert(i, P)
                mcdID += 1

    # start of Object 5 - Analog Inputs ID = 2 for all binary inputs MCD and SS or NL
    object5 = etree.Element("OBJECT", Id="2", Key="1300")
    object5.text = "\n"
    object5.tail = "\n"
    device2.insert(3, object5)

    objGroupNums = 1900
    k = 0
    h = 0
    j = 0
    # loops for creating the client side analog points based on their name and frame #
    # creating each frame or "GROUP"

    for i in range(analogPoints.__len__() / 4):

        analogGroups = etree.Element("GROUP", Id=str(i), Key=str(objGroupNums))
        analogGroups.text = "\n"
        analogGroups.tail = "\n"
        object5.insert(h, analogGroups)
        objGroupNums += 1

        for h in range(4):
            if analogPoints[k].get("description") == "UNDEFINED":
                P = etree.Element("P", Id=str(h))
                P.text = ""
                P.tail = "\n"
                analogGroups.insert(h, P)
            else:
                P = etree.Element("P", Id=str(h), Key=str(analogNums[j]), Name= str(analogPoints[k].get("description")))
                P.text = ""
                P.tail = "\n"
                analogGroups.insert(h, P)
            k += 1
            j += 1

    # start of Object 6 - Control Inputs ID = 4 for all Control inputs MCD and SS or NL
    object6 = etree.Element("OBJECT", Id="4", Key=str(objGroupNums))
    object6.text = "\n"
    object6.tail = "\n"
    device2.insert(4, object6)
    objGroupNums += 1

    # start of Object 5 - Analog Inputs ID = 2 for all binary inputs MCD and SS or NL
    type = etree.Element("Type")
    type.text = "1"
    type.tail = "\n"
    object6.insert(0, type)

    a = controlPoints.__len__() // 16
    e = controlPoints.__len__() % 16
    controlNumGroups = a
    if e != 0:
        controlNumGroups += 1

    a = 3
    h = 0
    j = 0
    k = 0
    for i in range(controlNumGroups):
        controlGroups = etree.Element("GROUP", Id=str(i + 1), Key=str(objGroupNums))
        controlGroups.text = "\n"
        controlGroups.tail = "\n"
        object6.insert(h+1, controlGroups)
        objGroupNums += 1

        for h in range(16):
            if controlPoints[k].get("description") == "UNDEFINED":
                P = etree.Element("P", Id=str(h))
                P.text = ""
                P.tail = "\n"
                controlGroups.insert(h, P)
            elif controlPoints[k].get("description") == "SPARE":
                P = etree.Element("P", Id=str(h), Key=str(controlNums[j]), Name=controlPoints[k].get("description"))
                P.text = ""
                P.tail = "\n"
                controlGroups.insert(h, P)
            elif controlPoints[k].get("description") != "" and controlPoints[k].get("function") == "":
                P = etree.Element("P", Id=str(h), Key=str(controlNums[j]), Name=controlPoints[k].get("description"))
                P.text = ""
                P.tail = "\n"
                controlGroups.insert(h, P)
            elif controlPoints[k].get("description") != "" and controlPoints[k].get("function") != "":
                P = etree.Element("P", Id=str(h), Key=str(controlNums[j]), Name=controlPoints[k].get("description") + "_" + str(sheet.cell_value(a, 12)).strip().replace(" ", "_")
                     .replace("\\", "_").replace("/", "_").replace("-", "_")
                     .replace("&", "").replace("<", "").replace(">", "")
                     .replace('"', "").upper())
                P.text = ""
                P.tail = "\n"
                controlGroups.insert(h, P)
            if sheet.cell_value(a, 12).upper() == "DISABLE":
                invert = etree.Element("Invert")
                invert.text = "0"
                invert.tail = "\n"
                P.insert(0, invert)
                P.text = "\n"
            a += 1
            k += 1
            j += 1

    # start of Object 7 - Binary Inputs ID = 4 for all Control inputs MCD and SS or NL
    object7 = etree.Element("OBJECT", Id="7", Key=str(objGroupNums))
    object7.text = "\n"
    object7.tail = "\n"
    device2.insert(5, object7)
    objGroupNums += 1

    try:
     ssID = int(ssPoints[0].get('index'))
    except IndexError:
        print("")

    # add back in if the there need to be a conditional for going over 55
    # if ssPoints.__len__() > 55:

    # start of Object 4 - Binary Inputs first ID = 1 for MCD on the Client side
    group1 = etree.Element("GROUP", Id="0", Key=str(objGroupNums))
    group1.text = "\n"
    group1.tail = "\n"
    object7.insert(0, group1)
    objGroupNums += 1

    for i in range(ssPoints.__len__()):
        if ssPoints[i].get("description") == "SPT_COMM_FAIL":
            print("")
            ssID += 1
        else:
            if ssPoints[i].get('description').upper() != "UNDEFINED":
                P = etree.Element("P", Id=str(ssID), Key=str(ssBinaryNums[i]), Name=ssPoints[i].get('description'))
                P.text = ""
                P.tail = "\n"
                group1.insert(i, P)
                ssID += 1
            else:
                P = etree.Element("P", Id=str(ssID))
                P.text = ""
                P.tail = "\n"
                group1.insert(i, P)
                ssID += 1

        # add in if another group needs to be created for when the simple status points go over 55 points
        # group2 = etree.Element("GROUP", Id="1", Key=str(objGroupNums))
        # group2.text = "\n"
        # group2.tail = "\n"
        # object7.insert(1, group2)
        # objGroupNums += 1
        #
        # for i in range(55, ssPoints.__len__()):
        #
        #     if ssPoints[i].get('description').upper() != "UNDEFINED":
        #         if ssPoints[i].get("description") == "SPT_COMM_FAIL":
        #             print("")
        #             ssID += 1
        #         else:
        #             P = etree.Element("P", Id=str(ssID), Key=str(ssBinaryNums[i]), Name=ssPoints[i].get('description'))
        #             P.text = ""
        #             P.tail = "\n"
        #             group2.insert(i, P)
        #             ssID += 1
        #     else:
        #         P = etree.Element("P", Id=str(ssID))
        #         P.text = ""
        #         P.tail = "\n"
        #         group2.insert(i, P)
        #         ssID += 1

    # start of Object 8 - Binary Inputs ID = 4 for all Control inputs MCD and SS or NL
    object8 = etree.Element("OBJECT", Id="16384", Key=str(objGroupNums))
    object8.text = "\n"
    object8.tail = "\n"
    device2.insert(6, object8)

    for i in range(ssPoints.__len__()):
        if ssPoints[i].get("description") == "SPT_COMM_FAIL":
            P8 = etree.Element("P", Id="0", Key=str(ssBinaryNums[i]), Name=ssPoints[i].get("description"))
            P8.text = ""
            P8.tail = "\n"
            object8.insert(0, P8)

    for i in range(mcdPoints.__len__()):
        if mcdPoints[i].get("description") == "SPT_COMM_FAIL":
            P8 = etree.Element("P", Id="0", Key=str(mcdBinaryNums[i]), Name=mcdPoints[i].get("description"))
            P8.text = ""
            P8.tail = "\n"
            object8.insert(0, P8)

    tree = etree.ElementTree(root)
    tree.write(str(stationName) + " ASE SPT Alert 9550 DNP Rev " + str(revision) + ".xml")
    # add for addtional info - - - > ,xml_declaration=True,   encoding="utf-8")

# arguments are rxTxUserInput, stationName, revision, emsBaudRate, rtuBaudRate, emsLocalAddress, rtuRemoteAddress, enableUmode, port#.
trw9550Gen("YES", "Moore", "A", "9600", "2400", "10191", "3", "No", "2")
