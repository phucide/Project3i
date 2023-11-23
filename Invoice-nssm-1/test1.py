import json
with open("SelectedMonths.json", 'r', encoding='utf-8') as file:
        data = json.load(file)
month_from = int(data["MonthFrom"])
month_to = int(data["MonthTo"])


while month_from<=month_to:
    print(month_from)
    month_from+=1

