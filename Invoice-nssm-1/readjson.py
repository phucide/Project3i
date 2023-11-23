import json
import xml.etree.ElementTree as ET
xml_file = 'Invoice_Setting.xml'
tree = ET.parse(xml_file)
root = tree.getroot()
variable = root.find("variable")

with open(variable.find("json_sel_month").text, 'r', encoding='utf-8') as file:
        data = json.load(file)
if int(data["MonthFrom"])<10:
    from_date = "01/0"+data["MonthFrom"]+"/2023"
else:
    from_date = "01/"+data["MonthFrom"]+"/2023"

print(from_date)
