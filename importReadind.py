import xml.etree.ElementTree as ET
import requests
import time
from openpyxl import Workbook
def xml_to_dict(element):
    data = {}
    for child in element:
        data[child.tag] = child.text
    return data

def print_xml_dict(data):
    for key, value in data.items():
        print(key, ":", value)

xml_file = 'data1.xml'

def fillsheet(sheet,data,row):
    for column, value in enumerate(data,1):
        sheet.cell(row = row, column = column, value = value)

def returnStrDayList(startYear, startMonth, endYear, endMonth, day = "01"):
    result = []
    if startYear == endYear:
        for month in range(startMonth, endMonth+1):
            month = str(month)
            if len(month) == 1:
                month = "0" + month
            result.append(str(startYear)+month+day)    
    for year in range(startYear, endYear+1):
        if year == startYear:
            for month in range(startMonth, 13):
                month = str(month)
                if len(month) == 1:
                    month = "0" + month
                result.append(str(year)+month+day)    
        elif year == endYear:
            for month in range(1, endMonth+1):
                month = str(month)
                if len(month) == 1:
                    month = "0" + month
                result.append(str(year)+month+day)                
        else:
            for month in range(1, 13):
                month = str(month)
                if len(month) == 1:
                    month = "0" + month
                result.append(str(year)+month+day)
        return result            

tree = ET.parse(xml_file)
root = tree.getroot()
xml_dict = xml_to_dict(root)
print_xml_dict(xml_dict) #原來xml_dict是字典
fields = ["Date","Trade Volume",
"Trade Value","Opening Price",
"Highest Price","Lowest Price",
"Closing Price","Change","Transaction"]
wb =Workbook()
sheet=wb.active 
sheet.title = "fields"
fillsheet(sheet, fields, 1)
startyear, startmonth, = int(xml_dict["startYear"]), int(xml_dict["startMonth"])
endYear, endMonth = int(xml_dict["endYear"]), int(xml_dict["endMonth"])

yearlist = returnStrDayList(startyear, startmonth, endYear, endMonth)
print(yearlist) 

row = 2

for yearMonth in yearlist:
    rq = requests.get(xml_dict["url"], params={
        "response":"json",
        "date": yearMonth,
        "stockNo": xml_dict["stockNo"]
    })
jsonData = rq.json()
dailyPriceList = jsonData.get("data", [])
for dailyPrice in dailyPriceList:
    fillsheet(sheet, dailyPrice, row)
    row += 1
time.sleep(3)    
name = xml_dict["excelName"]
wb.save("idk.xlsx") #saving






