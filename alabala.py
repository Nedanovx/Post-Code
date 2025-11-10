import re
import requests
from pprint import pprint
from openpyxl import Workbook
addresses = [
    "С.АБЛАНИЦА ОбЩ.ХАДЖИДИМОВО ОБЛ.БЛАГОЕВГРАД ул. Център 5", #2932
    "С.АБЛАНИЦА ОБЩ.ЛОВЕЧ ОБЛ.ЛОВЕЧ ул. Странджа 11", #5574
    "С.АБЛАНИЦА ОБЩ.ПАЗАРДЖИК ОБЛ.ВЕЛИНГРАД ул. Липа 3",#грешни данни/разменени община и област
    "С.СТАРА РЕКА ОБЩ.тунджа ОБЛ.ямбол ул. Георги Кирков 44", #8675
    "С.НОВА МАХАЛА; ОБЩ.ПЛОВДИВ ОБЛ.ЛЪКИ ул. Рила 18", #грешни данни има ; но регекса не го хваща, общ и обл са измислени
    "С.НОВА ВАСИЛЕВА ОБЩ.СТАРА ЗАГОРА ОБЛ.ОПАН ул. Лале 7",#грешни данни / няма нова василева и len(data)==0
    "С.СТАРА РЕКА ОБЩ.Сливен ОБЛ.СЛИВЕН ул. Дюлева 9", #8841
    "С.СТАРА ЗАГОРА ОБЩ.СТАРА ЗАГОРА ОБЛ.СТАРА ЗАГОРА ул. Момина Баня 10", #няма такова село но в бг пощи се 
    #подава само името - въпроса е дали това са грешни данни и ако са как да ги прескоча
    "ГР.СТАРА ЗАГОРА ОБЩ.СТАРА ЗАГОРА ОБЛ.СТАРА ЗАГОРА бул. Цар Симеон Велики 103",#6000
    "ГР.НОВА ЗАГОРА ОБЩ.СЛИВЕН ОБЛ.НОВА ЗАГОРА бул. Христо Ботев 2",#8900
    "ГР.БАЛЧИК ОБЩ.ДОБРИЧ ОБЛ.БАЛЧИК ул. Кубрат 1",#9600
    "С.НОВИ ИСКЪР ОБЩ.СОФИЯ ОБЛ.СТОЛИЧНА ул. 1-ва 3", #1280
    "С.Победа ОБЩ.Добрич ОБЛ.Добрич ул. Втора 5" # в бг пощи общината се казва Добрич-селска , но в адреса е само Добрич????
]
wb = Workbook()
ws = wb.active
ws.title = "Postal Codes"
ws.append(["Адрес", "Пощенски код"])
count=0
pattern = re.compile(
    r'(?i)\b(?P<type>ГР|С)\.\s*'
    r'(?P<city>[А-Яа-я]+(?:\s[А-Яа-я]+)*)\s*'
    r'(?:[;,\s]+)'
    r'ОБЩ\.\s*'
    r'(?P<municipality>[А-Яа-я]+(?:\s[А-Яа-я]+)*)\s*'
    r'(?:[;,\s]+)'
    r'ОБЛ\.\s*'
    r'(?P<region>[А-Яа-я]+(?:\s[А-Яа-я]+)*?)'
    r'(?=\s+(ул\.|бул\.|\d|$))'
)
url = "https://bgpostcode.com/api/v1/city?name="
for address in addresses:
    count+=1
    postcode =''
    #print(address)
    match = pattern.search(address)
    # print(match)
    if match:
        type= match.group(1)
        city = match.group(2)
        obshtina = match.group(3)
        oblast = match.group(4)
        # print(f"Type: {type}, City: {city}, Region: {oblast}, Municipality: {obshtina}")
        # print(f'{url}{city}')
        response = requests.get(f'{url}{city}')
        if response.status_code == 200:
            data = response.json()
            if len(data) == 0:
                print(f'Count {count} not found {city}')
                pass
            # print(f'Postcode: {type.lower()}')
            if(type.lower() =='гр'):
                postcode = data[0]['postcode']
                print(f'Count {count} City {city} Postcode: {postcode}')
            else:
                for item in range(len(data)):
        # pprint(item)
        # pprint(data[item]['municipality'])
                    
                    if len(data)==1:
                        postcode = data[0]['postcode']
                        print(f'Count {count} City {city} Postcode: {postcode}')
                        break
                    elif data[item]['municipality']['name'].lower() == obshtina.lower() and data[item]['region']['name'].lower() == oblast.lower():
                        postcode = data[item]['postcode']
                        print(f'{count} City {city} Postcode: {postcode}')
                        break    
                if postcode =='':
                    print(f'Count {count} not found {city} {obshtina} {oblast}')
                    postcode ='' 
            ws.append([address, postcode])                            
        else:
            print(f'{response.status_code}')
            ws.append([address, ''])
    else:
        print(f'No match')                        
wb.save("postal_codes.xlsx")
print("done")                    
