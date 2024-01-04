import requests
import schedule
from bs4 import BeautifulSoup
import fake_useragent
import openpyxl



def parsing_LI():
    sites_list_number = 0
    stats_list_number = 0
    url = "https://www.liveinternet.ru/rating/ru/sport/today.tsv?"

    fua = fake_useragent.UserAgent().random
    header = {
        'user-agent': fua
    }

    resource = requests.get(url).text

    sport_list = resource.split("\t")

    STEP_SITES_IN_FILE_FROM_LI = 5
    sport_sites = [] # create list sports sites

    for b in sport_list:
        if STEP_SITES_IN_FILE_FROM_LI <= len(sport_list):
            sport_sites.append(sport_list[c])
            STEP_SITES_IN_FILE_FROM_LI += 6

    stat_list = [] # site's statistics per day
    for word in sport_list:
        if word.isnumeric():
            stat_list.append(int(word))

    STEP_VALUES_SITES_FROM_LI=2

    summers = []

    for i in stat_list:
        if STEP_VALUES_SITES_FROM_LI < len(stat_list):
            summers.append(stat_list[a])
            STEP_VALUES_SITES_FROM_LI += 2


    book = openpyxl.Workbook()

    sheet = book.active

    sheet.column_dimensions['A'].width = 50
    sheet.column_dimensions['B'].width = 20

    # cells_sites = [] # create list cells in excel file
    #
    # for i in sport_sites: # добавление ячеек
    #     cells_sites.append(f'A{sport_sites.index(i) + 1}')
    #
    #
    # for cell in cells_sites: # добавление в ячейки файла названий сайтов
    #     sheet[cell] = sport_sites[cells_sites.index(cell)]
    #
    # cells_stats = [] # create list cells in excel file
    #
    # for i in summers: # добавление ячеек
    #     cells_stats.append(f'B{summers.index(i) + 1}')
    #
    # for cell in cells_stats: # добавление в ячейки файла названий сайтов
    #     sheet[cell] = summers[cells_stats.index(cell)]

    cells_sites = []  # create list cells in excel file

    for i in sport_sites:  # добавление ячеек
        cells_sites.append(f'A{sites_list_number + 1}')
        sites_list_number +=1
    print(cells_sites)

    for cell in cells_sites:  # добавление в ячейки файла названий сайтов
        sheet[cell] = sport_sites[cells_sites.index(cell)]

    cells_stats = []  # create list cells in excel file

    for i in summers:  # добавление ячеек
        cells_stats.append(f'B{stats_list_number + 1}')
        stats_list_number += 1
    print(cells_stats)

    for cell in cells_stats:  # добавление в ячейки файла статистики сайтов
        sheet[cell] = summers[cells_stats.index(cell)]

    stats_list_number += 30
    sites_list_number += 30


    book.save("massive.xlsx")
    book.close()

def main():

    schedule.every(5).seconds.do(parsing_LI)

    while True:
        schedule.run_pending()

if __name__ == '__main__':
    main()

# results_day = dict(zip(sport_sites, summers))
#
# for i in results_day.items():
#     print(i[0], i[1])
#
# print (results_day)