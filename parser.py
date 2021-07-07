import json
import requests
import xlsxwriter

URL = 'https://sigame.xyz/'
HEADERS = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                         'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.101 Safari/537.36',
           'accept': '*/*'}
URL_PAC = 'api/pack?sort=created_at&search=all&'

data_list = []


def get_content(req):
    data = json.loads(req)
    for i in data['data']:
        count = 0
        list_of_rounds = []
        list_of_themes = []
        for j in i['json_contents']['rounds']:
            count += 1
            list_of_rounds.append(j['name'])
            for k in j['themes']:
                list_of_themes.append(k['name'])
        data_list.append(
            {
                "Название": i['json_contents']['name'],
                "Количество раундов": count,
                "Список раундов": list_of_rounds,
                "Список тем": list_of_themes,
                "Ссылка для скачивания": URL + 'api/pack/' + str(i['id']) + '/download'
            }
        )


def parse():
    for item in range(1, 41):
        req = requests.get(URL + URL_PAC + f"page={item}&per_page=8", HEADERS)
        if req.status_code == 200:
            get_content(req.text)
        else:
            break
    write_in_table()


def write_in_table():
    with xlsxwriter.Workbook('tablee.xlsx') as workbook:
        worksheet = workbook.add_worksheet()
        worksheet.write_row(0, 0, ["Название", "Количество раундов", "Список раундов", "Список тем", "Ссылка для скачивания"])
        count = 1
        for data in data_list:
            rounds = ''
            for elem in data['Список раундов']:
                rounds = rounds + elem + '; '
            themes = ''
            for th in data['Список тем']:
                themes = themes + th + '; '
            p = [data['Название'], str(data["Количество раундов"]), rounds, themes, data["Ссылка для скачивания"]]
            worksheet.write_row(count, 0, p)
            count += 1


parse()
