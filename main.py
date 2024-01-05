import requests
import schedule
from bs4 import BeautifulSoup
import fake_useragent
import openpyxl

def main():

    summers = 0

    parsing_LI()
    create_list_name_sportsites()
    create_list_values_sportssites()
    push_values_in_excel()
    get_timeout()
def parsing_LI(): # функция для парсинга

    url = "https://www.liveinternet.ru/rating/ru/sport/today.tsv?"

    fua = fake_useragent.UserAgent().random
    header = {
        'user-agent': fua
    }

    resource = requests.get(url).text
    global sport_list
    sport_list = resource.split("\t")

def create_list_name_sportsites(): # для создания списка сайтов

    STEP_SITES_IN_FILE_FROM_LI = 5
    global sport_sites
    sport_sites = [] # create list sports sites

    for b in sport_list:
        if STEP_SITES_IN_FILE_FROM_LI <= len(sport_list):
            sport_sites.append(sport_list[STEP_SITES_IN_FILE_FROM_LI])
            STEP_SITES_IN_FILE_FROM_LI += 6

def create_list_values_sportssites():
    global stat_list
    stat_list = [] # site's statistics per day
    for word in sport_list:
        if word.isnumeric():
            stat_list.append(int(word))

    STEP_VALUES_SITES_FROM_LI = 2
    global summers
    summers = []
    for i in stat_list:
        if STEP_VALUES_SITES_FROM_LI < len(stat_list):
            summers.append(stat_list[STEP_VALUES_SITES_FROM_LI])
            STEP_VALUES_SITES_FROM_LI += 2

    print()

def push_values_in_excel():

    sites_list_number = 0
    stats_list_number = 0

    book = openpyxl.Workbook()
    book.remove(book.active)
    sheet_1 = book.create_sheet("Данные")

    sheet_1.column_dimensions['A'].width = 50
    sheet_1.column_dimensions['B'].width = 20

    cells_sites = []  # create list cells in excel file

    for i in sport_sites:  # добавление ячеек
        cells_sites.append(f'A{sites_list_number + 1}')
        sites_list_number +=1
    print(cells_sites)

    for cell in cells_sites:  # добавление в ячейки файла названий сайтов
        sheet_1[cell] = sport_sites[cells_sites.index(cell)]

    cells_stats = []  # create list cells in excel file

    for i in summers:  # добавление ячеек
        cells_stats.append(f'B{stats_list_number + 1}')
        stats_list_number += 1
    print(cells_stats)

    for cell in cells_stats:  # добавление в ячейки файла статистики сайтов
        sheet_1[cell] = summers[cells_stats.index(cell)]




    book.save("massive.xlsx")
    book.close()

def get_timeout():

    schedule.every(5).seconds.do(parsing_LI)

    while True:
        schedule.run_pending()

if __name__ == '__main__':
    main()