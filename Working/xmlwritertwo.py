from lxml import etree

root = etree.Element("SPT", ID = "1", Ref = "18")
aEL = etree.Element("NewEL", ID = "2")
aEL.text = "4" 


#root.set('Ref','18')

a = etree.Element("a")
a.text = "1"
a.append(aEL)
root.append(a)
b = etree.Element("b")
b.text = "2"
root.insert(1,b)
c = etree.Element("c")
c.text = "3"
root.insert(2,c)


tree = etree.ElementTree(root)
tree.write("sample.xml",pretty_print=True)



#add for addtional info - - - > ,xml_declaration=True,   encoding="utf-8")  



#tree = etree.ElementTree(ROOT)
#tree.write("filename2.xml", pretty_print=True)
