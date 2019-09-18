from lxml import etree
import xlrd


file_loc = ("C:\Users\wblankenship\Documents\GitHub\SPT-Builder\SPT Config Generator.xlsx")
wb = xlrd.open_workbook(file_loc)
sheet = wb.sheet_by_name("Sheet1")
rows = sheet.nrows



def s9000Gen(invertRXTX):

    # root of the document
    root = etree.Element("SPT", Id="1")
    root.text = "\n"

    blank = etree.Element("blank", Id = "0")

    mcdBinaryTags = [None] * sheet.nrows
    ssBinaryTags = [None] * sheet.nrows
    controlTags = [None] * sheet.nrows
    analogTags = [None] * sheet.nrows
    mcdNums = 0;
    ssNums = 0;
    controlNums = 0;
    analogNums = 0

    # creating the variables for number of rows with non blank cells
    for i in range(3, sheet.nrows):
        if sheet.cell_value(i, 3) != "":
            mcdNums += 1
        if sheet.cell_value(i, 7) != "":
            ssNums += 1
        if sheet.cell_value(i, 11) != "":
            controlNums += 1
        if sheet.cell_value(i, 16) != "":
            analogNums += 1

    # for all tags with text and tail as \n & for all settings tags with one piece of text
    def tagGen1(dict):

            tagName = str(dict.get("tagName"))
            Idee = str(dict.get("id"))
            insTag = dict.get("insTag")
            text = dict.get("text")
            tail = dict.get("tail")
            insNum = dict.get("insNum")
            attName = dict.get("Name")
            key = dict.get("key")

            # checking for all the different cases of tag creation
            if Idee != "" and attName != "" and key != "":
                tag = etree.Element(tagName, Id=Idee, Key=key, Name=attName)
                tag.text = text
                tag.tail = tail
                insTag.insert(insNum, tag)
            elif Idee != "" and key != "":
                tag = etree.Element(tagName, Id=Idee, Key=key)
                tag.text = text
                tag.tail = tail
                insTag.insert(insNum, tag)
            elif Idee != "" and attName != "":
                tag = etree.Element(tagName, Id=Idee, Name=attName)
                tag.text = text
                tag.tail = tail
                insTag.insert(insNum, tag)
            elif Idee != "":
                tag = etree.Element(tagName, Id=Idee)
                tag.text = text
                tag.tail = tail
                insTag.insert(insNum, tag)
            else:
                tag = etree.Element(tagName)
                tag.text = text
                tag.tail = tail
                insTag.insert(insNum, tag)
            j = 118
            if(tagName == "OBJECT" and Idee == "1" and key == ""):
                for i in range(mcdNums):
                    if str(sheet.cell_value(3+i,3)) != "":
                        mcdBinaryTags[i] = {"tagName": "P", "key": j, "ref": j, "frame" : sheet.cell_value(3+i,1), "index": int(sheet.cell_value(3+i,2)),
                                            "description": str(sheet.cell_value(3+i,3).strip().replace(" ", "_")
                                       .replace("\\","_").replace("/","_").replace("-","_")
                                       .replace("&", "").replace("<", "").replace(">", "")
                                       .replace('"',"").upper()), "insTag": tag, "text": "", "tail": "\n"}
                        tagGen2(mcdBinaryTags[i])
                    j+=1
                for i in range(ssNums):
                    if str(sheet.cell_value(3+i,7)) != "":
                        ssBinaryTags[i] = {"tagName": "P", "key": j, "ref": j, "frame" : sheet.cell_value(3+i,5), "index": int(sheet.cell_value(3+i,6)),
                                            "description": str(sheet.cell_value(3+i,7).strip().replace(" ", "_")
                                       .replace("\\","_").replace("/","_").replace("-","_")
                                       .replace("&", "").replace("<", "").replace(">", "")
                                       .replace('"',"").upper()), "insTag": tag, "text": "", "tail": "\n"}
                        tagGen2(ssBinaryTags[i])
                    j+=1
            if (tagName == "GROUP" and Idee == "0" and key == "117" or key == "197"):
                    for i in range(mcdNums):
                        if str(sheet.cell_value(3 + i, 3)) != "":
                            if key == "117":
                                mcdBinaryTags[i] = {"tagName": "P", "key": j, "ref": 0, "frame": sheet.cell_value(3 + i, 1),
                                                    "index": int(sheet.cell_value(3 + i, 2)),
                                                    "description": str(sheet.cell_value(3 + i, 3).strip().replace(" ", "_")
                                                                       .replace("\\", "_").replace("/", "_").replace("-", "_")
                                                                       .replace("&", "").replace("<", "").replace(">", "")
                                                                       .replace('"', "").upper()), "insTag": tag, "text": "", "tail": "\n"}
                                tagGen2(mcdBinaryTags[i])
                        j+=1

                    for i in range(ssNums):
                        if str(sheet.cell_value(3+i,7)) != "" :
                            if key == "197":
                                ssBinaryTags[i] = {"tagName": "P", "key": j, "ref": 0, "frame" : sheet.cell_value(3+i,5), "index": int(sheet.cell_value(3+i,6)),
                                                    "description": str(sheet.cell_value(3+i,7).strip().replace(" ", "_")
                                               .replace("\\","_").replace("/","_").replace("-","_")
                                               .replace("&", "").replace("<", "").replace(">", "")
                                               .replace('"',"").upper()), "insTag": tag, "text": "", "tail": "\n"}
                                tagGen2(ssBinaryTags[i])
                        j+=1

            if (tagName == "OBJECT" and Idee == "1" and key == ""):
                for i in range(mcdNums):
                    if str(sheet.cell_value(3 + i, 3)) != "":
                        mcdBinaryTags[i] = {"tagName": "P", "key": j, "ref": j, "frame": sheet.cell_value(3 + i, 1),
                                            "index": int(sheet.cell_value(3 + i, 2)),
                                            "description": str(sheet.cell_value(3 + i, 3).strip().replace(" ", "_")
                                                               .replace("\\", "_").replace("/", "_").replace("-", "_")
                                                               .replace("&", "").replace("<", "").replace(">", "")
                                                               .replace('"', "").upper()), "insTag": tag, "text": "",
                                            "tail": "\n"}
                        tagGen2(mcdBinaryTags[i])
                    j += 1


                # for i in range(ssNums):
                #     if str(sheet.cell_value(3 + i, 7)) != "":
                #         mcdBinaryTags[i] = {"tagName": "P", "key": j, "ref": j, "frame": sheet.cell_value(3 + i, 5),
                #                             "index": int(sheet.cell_value(3 + i, 6)),
                #                             "description": str(sheet.cell_value(3 + i, 7).strip().replace(" ", "_")
                #                                                .replace("\\", "_").replace("/", "_").replace("-", "_")
                #                                                .replace("&", "").replace("<", "").replace(">", "")
                #                                                .replace('"', "").upper()), "insTag": tag, "text": "",
                #                             "tail": "\n"}
                #         tagGen2(mcdBinaryTags[i])
                #     j += 1

            return tag

    def tagGen2(dict):

            tagName = str(dict.get("tagName"))
            Idee = str(dict.get("index"))
            insTag = dict.get("insTag")
            text = dict.get("text")
            tail = dict.get("tail")
            insNum = dict.get("index")
            description = str(dict.get("description"))
            key = str(dict.get("key"))
            ref = str(dict.get("ref"))

            # checking for all the different cases of tag creation

            if Idee != "" and description != "" and ref == "0":
                tag = etree.Element(tagName, Id = Idee, Key=key, Name=description)
                tag.text = text
                tag.tail = tail
                insTag.insert(insNum, tag)
            elif Idee != "" and ref != "":
                tag = etree.Element(tagName, Id=Idee, Key=key)
                tag.text = text
                tag.tail = tail
                insTag.insert(insNum, tag)




    #print(mcdBinaryTags[10].get("description"))

    #Key:Value Pairs
    #Start of server side to EMS
    tag1 = {"tagName": "DIRECTION", "id": 0, "key": "", "Name":"", "insTag": root, "text": "\n", "tail": "\n", "insNum": 0}
    tag2 = {"tagName": "PROTOCOL", "id": 1, "key": "", "Name":"", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 0}
    tag3 = {"tagName": "LINE", "id": 1, "key": "", "Name":"", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 0}
    tag4 = {"tagName": "Interface", "id": "", "key": "", "Name":"", "insTag": blank, "text": "0", "tail": "\n", "insNum": 0}
    tag5 = {"tagName": "BaudRate", "id": "", "key": "", "Name":"", "insTag": blank, "text": "9600", "tail": "\n", "insNum": 1}
    tag6 = {"tagName": "SwitchedRX", "id": "", "key": "", "Name":"", "insTag": blank, "text": "0", "tail": "\n", "insNum": 2}
    tag7 = {"tagName": "SwitchedTX", "id": "", "key": "", "Name":"", "insTag": blank, "text": "0", "tail": "\n", "insNum": 3}
    tag8 = {"tagName": "DEVICE", "id": 10237, "key": "", "Name":"toEMS", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 4}
    # Binary MCD & SS points
    tag9 = {"tagName": "OBJECT", "id": 1, "key": "", "Name": "", "insTag": blank, "text": "\n", "tail": "\n","insNum": 0}


    # for elm, tag in enumerate(mcdBinaryTags):
    #     if tag != None:



    # Control Points
    tag10 = {"tagName": "OBJECT", "id": 12, "key": "", "Name": "", "insTag": blank, "text": "\n", "tail": "\n","insNum": 1}
    # Status Points
    tag11 = {"tagName": "OBJECT", "id": 30, "key": "", "Name": "", "insTag": blank, "text": "\n", "tail": "\n","insNum": 2}
    tag12 = {"tagName": "DIRECTION", "id": 1, "key": "111", "Name": "", "insTag": root, "text": "\n", "tail": "\n","insNum": 1}
    tag13 = {"tagName": "PROTOCOL", "id": 9, "key": "112", "Name": "", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 0}
    tag14 = {"tagName": "LINE", "id": 2, "key": "113", "Name":"", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 0}
    tag15 = {"tagName": "Interface", "id": "", "key": "", "Name": "", "insTag": blank, "text": "0", "tail": "\n", "insNum": 0}
    tag16 = {"tagName": "BaudRate", "id": "", "key": "", "Name": "", "insTag": blank, "text": "1200", "tail": "\n", "insNum": 1}

    # for RX TX inversion
    if (invertRXTX == True):
        tag17 = {"tagName": "InvertRX", "id": "", "key": "", "Name": "", "insTag": blank, "text": "1", "tail": "\n", "insNum": 2}
        tag18 = {"tagName": "InvertTX", "id": "", "key": "", "Name": "", "insTag": blank, "text": "1", "tail": "\n", "insNum": 3}
        tag19 = {"tagName": "DEVICE", "id": 3, "key": "114", "Name":"Millhurst_TRWS9", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 4}
        tag20 = {"tagName": "REQUEST", "id": 0, "key": "115", "Name": "", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 0}
        tag21 = {"tagName": "Function", "id": "", "key": "", "Name": "", "insTag": blank, "text": "5", "tail": "\n", "insNum": 0}
        tag22 = {"tagName": "Variation", "id": "", "key": "", "Name": "", "insTag": blank, "text": "1", "tail": "\n", "insNum": 1}
        tag23 = {"tagName": "XCount", "id": "", "key": "", "Name": "", "insTag": blank, "text": "1", "tail": "\n", "insNum": 2}
        tag24 = {"tagName": "YCount", "id": "", "key": "", "Name": "", "insTag": blank, "text": "3", "tail": "\n", "insNum": 3}
        # binary MCD and 2B
        tag25 = {"tagName": "OBJECT", "id": 1, "key": "116", "Name": "", "insTag": "request", "text": "\n", "tail": "\n", "insNum": 1}
        tag26 = {"tagName": "GROUP", "id": 0, "key": "117", "Name": "", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 0}
        # analog
        tag27 = {"tagName": "OBJECT", "id": 2, "key": "190", "Name": "", "insTag": "request", "text": "\n", "tail": "\n", "insNum": 2}
        tag28 = {"tagName": "GROUP", "id": 0, "key": "191", "Name": "", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 0}
        tag29 = {"tagName": "GROUP", "id": 1, "key": "192", "Name": "", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 1}
        tag30 = {"tagName": "GROUP", "id": 2, "key": "193", "Name": "", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 2}
        # control
        tag31 = {"tagName": "OBJECT", "id": 4, "key": "194", "Name": "", "insTag": "request", "text": "\n", "tail": "\n", "insNum": 3}

        remainder = controlNums % 16
        totalGroups = controlNums // 16
        if remainder != 0:
            totalGroups += 1
        for i in range(totalGroups):
            tag32 = {"tagName": "GROUP", "id": i+1, "key": str(195+i), "Name": "", "insTag": blank, "text": "\n", "tail": "\n", "insNum": i}
            tagList.append(tag32)
            tag33 = {"tagName": "OBJECT", "id": 4, "key": str(196+i), "Name": "", "insTag": "request", "text": "\n", "tail": "\n", "insNum": 4}
            tag34 = {"tagName": "GROUP", "id": 0, "key": str(197+i), "Name": "", "insTag": blank, "text": "\n", "tail": "\n", "insNum": 0}
            tag35 = {"tagName": "OBJECT", "id": 16384, "key": str(198+i), "Name": "", "insTag": "request", "text": "\n", "tail": "\n", "insNum": 5}


        tagList = [tag1, tag2, tag3, tag4, tag5, tag6, tag7, tag8, tag9, tag10, tag11, tag12, tag13, tag14,
                   tag15, tag16, tag17, tag18, tag19, tag20, tag21, tag22, tag23, tag24, tag25, tag26, tag27,
                   tag28, tag29, tag30, tag31, tag33, tag34, tag35]

    #insert taglist
    insTagList = [None] * 1000
    i = 1
    for elm, tag in enumerate(tagList):

        if tag['insTag'] == root:
            insTagList[i] = tagGen1(tag)
        else:

            if tag['insNum'] == 0:
                    tag['insTag'] = insTagList[i]
                    i+=1
                    insTagList[i] = tagGen1(tag)
            else:
                # odd situation where object would not add after request on device.
                if tag['insTag'] == "request":
                    tag['insTag'] = insTagList[8]
                    i += 1
                    insTagList[i] = tagGen1(tag)
                else:
                    tag['insTag'] = insTagList[i-1]
                    insTagList[i] = tagGen1(tag)

    tree = etree.ElementTree(root)
    tree.write("Example.xml" )

s9000Gen(True)

# for i in dict.values():
#     print(str(i))
#     raw_input()
#
#
# # printing keys
# for i in biMcdPoints:
#     for b in i.keys():
#         print ("Key: " + str(b))
#
# #printing values
# for i in biMcdPoints:
#     for b in i.values():
#         print ("Value: " + str(b))
#
# #printing the pairs
# for i in biMcdPoints:
#     for b in i.items():
#         print ("Item: " + str(b))

# #single point
# bipoint1 = {'frame': 0, 'index': 0, 'description': "BK 1 VCB"}
# bipoint2 = {'frame': 0, 'index': 2, 'description': "47417 VCB"}
# bipoint3 = {'frame': 0, 'index': 1, 'description': "BT VCB"}
#
# biMcdPoints = [bipoint1, bipoint2, bipoint3]
# # sorting by index
# biMcdPoints = sorted(biMcdPoints, key=lambda k: k['index'])
