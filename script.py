import requests
from bs4 import BeautifulSoup
import xlsxwriter

version = "0.1"
row = 2
parsing = 1
page = 1

workbook = xlsxwriter.Workbook('out.xlsx') # Создать файл
worksheet = workbook.add_worksheet('out')
bold_and_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
center = workbook.add_format({'align': 'center'})
worksheet.write('A1', 'czDevelopment. version ' + str(version) + ". Special for kosharik By baran4ik.")
worksheet.write('A2', 'Название', bold_and_center)
worksheet.write('B2', 'Ссылка', bold_and_center)
worksheet.write('C2', 'SKU', bold_and_center)
worksheet.write('D2', 'Подходит', bold_and_center)
worksheet.write('E2', 'Preview', bold_and_center)
worksheet.set_row(0, 20)
worksheet.set_column(0, 0, 35)
worksheet.set_column(1, 1, 20)
worksheet.set_column(2, 2, 17)
worksheet.set_column(3, 3, 20)
worksheet.set_column(4, 4, 20)

def check(url):
    global page
    page = page + 1
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'lxml')
    items = soup.find('label', class_='input-sort')
    print(items)
    if items == "None":
        global parsing
        state = "Stop"
        parsing = 0
    else:
        state = "Continue"
        parsing = 1

    return state

def parse(url):
    global row
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'lxml')
    items = soup.find_all('div', class_='product-layout product-list col-xs-12')
    for n, i in enumerate(items, start=1):
        row = row + 1

        itemNAM = i.find("div", class_="caption").find("h4").find("a").text.strip()
        itemLNK = i.find("div", class_="caption").find("h4").find("a")['href']
        itemSKU = i.find("div", class_="mod").text.strip()
        itemPOD = i.find("div", class_="cats-block").find("p").text.strip()

        try:
            itemIMG = i.find("div", class_="image").find("a").find("img")["src"]
        except AttributeError:
            itemIMG = "no_img"


        worksheet.write("A"+str(row), itemNAM, center)
        worksheet.write("B"+str(row), itemLNK, center)
        worksheet.write("C"+str(row), itemSKU, center)
        worksheet.write("D"+str(row), itemPOD, center)
        worksheet.write("E"+str(row), itemIMG, center)

url = input("Ссылка на страницу: ")
parse(url)
workbook.close()
print("Данные обработаны и сохранены в файл out.xlsx")
