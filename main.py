import time
from xml.dom import minidom
import json

import xlsxwriter

from xml.etree import ElementTree as ET

# file = minidom.parse('data/data.xml')
#
# #use getElementsByTagName() to get tag
# models = file.getElementsByTagName('Report')
#
# elementProp_1 = file.getElementsByTagName('Prop')
#
#
#
# print(models.length)
# print(models[0].attributes['Title'].value)

import xml.dom.minidom


workbook = xlsxwriter.Workbook('DataExport/Data.xlsx')
worksheet = workbook.add_worksheet()


# parse the XML file
xml_doc = xml.dom.minidom.parse('data/data.xml')
# get the root element
root = xml_doc.documentElement
#print('Root is',root)

packages = xml_doc.getElementsByTagName('Report')

index_row_data = 3
index_colum_data = 0
#print(packages.length)
def beginHeaderFile():
    index_row_header = 0
    index_colum_header = 0
    element_1 = packages[0].getElementsByTagName('Prop')
    for i in element_1:
        # BEGIN HEADER EXCEL FILE
        getPropName = i.attributes["Name"].value
        worksheet.write(index_row_header, index_colum_header, getPropName)
        getPropLength = i.getElementsByTagName("Prop")
        if getPropLength.length != 0 and getPropLength.length != 88:
            for i in getPropLength:
                worksheet.write(index_row_header + 1, index_colum_header, i.attributes["Name"].value)
                index_colum_header += 1
        else:
            index_colum_header += 1
    workbook.close()

def formatFile():
    worksheet.merge_range("A1:B1")
    worksheet.merge_range("AG1:BF1")
    worksheet.merge_range("DG1:DI1")
    worksheet.merge_range("DN1:DR1")
    worksheet.merge_range("DX1:EC1")
    worksheet.merge_range("EJ1:EL1")
    workbook.close()



beginHeaderFile()
listData = []












for index in packages:
    element_1 = index.getElementsByTagName('Prop')
    for i in element_1:
        propLength = i.getElementsByTagName("Prop").length

        if propLength == 88: continue
        elif propLength == 0:
            elementIndex = i.getElementsByTagName("Value").length
            if elementIndex == 1:
                # value = i.getElementsByTagName("Value")[0].childNodes[0].data
                # if value is None: continue
                # else: print(value)
                print("AAA")
            else:
                for index in range(0, elementIndex - 1):
                    print(i.getElementsByTagName("Value")[index].childNodes[0])

        # else:
        #     print("aaa")

        # totalIndex = i.getElementsByTagName("Value").length
        # print(totalIndex)
        # if totalIndex != 1:
        #     for index in range(0, totalIndex):
        #         print(i.getElementsByTagName("Value")[index].childNodes[0].data)
        #         # listData.append(i.getElementsByTagName("Value")[index].childNodes[0].data)
        # else:
        #     print("-----1 element-------")
        #     print(i.getElementsByTagName("Value")[0].childNodes[0].data)
        #     print("-----1 element-------")
        #     # listData.append(i.getElementsByTagName("Value")[0].childNodes[0].data)



















        #data = i.getElementsByTagName("Value")[0].childNodes[0].data
        # dataLength = i.getElementsByTagName("EnabledExpression").length

        # print(dataLength)

        # if dataLength != 1 and dataLength != 0:
        #     print(i.getElementsByTagName("Name"))
        # elif dataLength == 1:
        #     print(i.getElementsByTagName("Value")[0].childNodes[0].data)






            #print(getProp[0].attributes['Name'].value)


    while True:
        time.sleep(1)