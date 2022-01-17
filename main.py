import requests, xlwt, datetime, os
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

CRM_URL = 'https://example.com'
CRM_LOGIN = 'somelogin'
CRM_PASS = 'somepassword'
CRM_API_KEY = 'someapikey'

today = datetime.date.today()


def last_day_of_month(date):
    if date.month == 12:
        return date.replace(year=date.year, day=31)
    return date.replace(month=date.month+1, day=1) - datetime.timedelta(days=1)


def get_teacher_data(list_tr):
    data = {
        "Teacher": "",
        "Groups": [

        ],
        "Hours": "",
        "Salary": ""
    }
    for tr in list_tr:
        tds = tr.find_all("td")
        if len(tds) > 6:
            teacher = tds[0].find("a").text
            data["Teacher"] = teacher
            for td in tds[1:]:
                if tds.index(td) == 1:
                    group = td.find("a").text
                elif tds.index(td) == 2:
                    days = td.find_all("x-ts-day")
                    days = [day.text.strip()[:5] for day in days]
                elif tds.index(td) == 3:
                    hours = td.text.replace(" а.ч.", "")
                elif tds.index(td) == 4:
                    stavka = td.text.replace("/астр.ч.", "")
                elif tds.index(td) == 5:
                    summ = td.text.replace("\xa0", "").replace(" злотых", "")
            data["Groups"].append({
                "Name": group,
                "Days": days,
                "Hours": hours,
                "Stavka": stavka,
                "Summ": summ
            })
        elif len(tds) > 4:
            for td in tds:
                if tds.index(td) == 0:
                    group = td.find("a").text
                elif tds.index(td) == 1:
                    days = td.find_all("x-ts-day")
                    days = [day.text.strip()[:5] for day in days]
                elif tds.index(td) == 2:
                    hours = td.text.replace(" а.ч.", "")
                elif tds.index(td) == 3:
                    stavka = td.text.replace("/астр.ч.", "")
                elif tds.index(td) == 4:
                    summ = td.text.replace("\xa0", "").replace(" злотых", "")
            data["Groups"].append({
                "Name": group,
                "Days": days,
                "Hours": hours,
                "Stavka": stavka,
                "Summ": summ
            })
        else:
            for td in tds:
                if tds.index(td) == 1:
                    data["Hours"] = td.text.replace(" а.ч.", " godz.astr.")
                elif tds.index(td) == 2:
                    data["Salary"] = td.text.replace("\xa0", "").replace(" злотых", " złotych")
    return data


def get_html(date_from, date_to, teachers):
    content = []
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    driver.get(CRM_URL)
    driver.set_window_size(1920, 1080)
    driver.find_element(By.ID, 'LogLogin').send_keys(CRM_LOGIN)
    driver.find_element(By.ID, 'LogPassword').send_keys(CRM_PASS)
    driver.find_element(By.XPATH, '//*[@id="cardMain"]/div/div/div[3]/form/div/div[4]/button').click()
    for teacher in teachers:
        driver.get(f"{CRM_URL}/Chart/TeachersStatistics?Submitted=True&Page=0&School=-1&BeginDate={date_from}"
                   f"T00%3A00%3A00&EndDate={date_to}T00%3A00%3A00&TeacherId={teacher}"
                   f"&DisciplineId=&MaturityId=&BeginMonth=&EndMonth=&LearningTypeId=&IsPartTime=false&IsNotPartTime"
                   f"=false&IsNativeSpeaker=False&IsNotNativeSpeaker=False")
        content.append(driver.page_source.encode('utf-8').strip())
    driver.close()
    return content


def write_data_to_file(data, filename, date):
    print(data)
    wb = xlwt.Workbook()
    sheet = wb.add_sheet("List1")
    xlwt.add_palette_colour("custom_colour", 0x21)
    wb.set_colour_RGB(0x21, 251, 228, 228)
    style = xlwt.easyxf('font: bold 0')
    sheet.write(1, 0, 'Ewidencja godzin wykonywania pracy\nzgodnie z umową zlecenia z dnia', style)
    sheet.write(1, 2, 'Miesiąc', style)
    sheet.write(1, 3, f'{date.strftime("%B")}, {today.year}', style)
    sheet.write(2, 0, 'Imię, nazwisko nauczyciela', style)
    sheet.write(2, 1, data["Teacher"], style)
    sheet.write(4, 0, '№', style)
    sheet.write(4, 1, 'Godz., godz.astr.', style)
    sheet.write(4, 2, 'Stawka, złotych/godz.astr.', style)
    sheet.write(4, 3, 'Opłata, złotych', style)
    sheet.write(4, 4, 'Dni', style)
    row = 5
    col = 0
    for group in data["Groups"]:
        sheet.write(row, col, data["Groups"].index(group) + 1, style)
        sheet.write(row, col + 1, group["Hours"], style)
        sheet.write(row, col + 2, group["Stavka"], style)
        sheet.write(row, col + 3, group["Summ"], style)
        for day in group["Days"]:
            sheet.write(row, col + 4 + group["Days"].index(day), day, style)
        row += 1
    sheet.write(row, 0, 'Razem', xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour; font: bold 1; borders: bottom dashed'))
    sheet.write(row, 1, data["Hours"], xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour; font: bold 1; borders: bottom dashed'))
    sheet.write(row, 2, "", xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour; font: bold 1; borders: bottom dashed'))
    sheet.write(row, 3, data["Salary"], xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour; font: bold 1; borders: bottom dashed'))
    row += 2
    sheet.write(row, 0, 'Podpis', style)
    sheet.write(row, 3, data["Teacher"], style)

    sheet.col(0).width = 256 * len('zgodnie z umową zlecenia z dnia')
    sheet.col(1).width = 256 * len(data["Teacher"])
    sheet.col(2).width = 256 * len(data["Teacher"])
    sheet.col(3).width = 256 * len('Оплаты преподавателю')
    dir = str(date.strftime("%B"))
    wb.save(f'{dir}/{filename}.xls')


def get_school_teachers(school_id):
    teachers_id = []
    authkey = CRM_API_KEY
    url = f"{CRM_URL}/Api/V2/GetTeachers"
    params = {
        "authkey": authkey,
        "officeOrCompanyId": school_id
    }
    response = requests.get(url, params=params)
    for item in response.json()["Teachers"]:
        if item.get("Status") == "Уволен":
            continue
        teachers_id.append(item.get("Id"))
    return teachers_id


teachers = get_school_teachers(1033)
print(teachers)
while True:
    user_answer = input("Введите номер месяца(1-12): ")
    if user_answer.isnumeric() and 12 >= int(user_answer) > 0:
        if len(user_answer) < 2:
            user_answer = f"0{user_answer}"
        break
    else:
        print("Неверный формат. Повторите")

year = today.year if today.month >= int(user_answer) else today.year - 1
date_from = f"{year}-{user_answer}-01"
date_to = last_day_of_month(datetime.date(year, int(user_answer), today.day))
path = os.getcwd()
dir = str(date_to.strftime("%B"))
path = os.path.join(path, dir)
if not os.path.exists(dir):
    os.mkdir(path)

content = get_html(date_from, date_to, teachers)
for item in content:
    soup = BeautifulSoup(item, "html.parser")
    table = soup.find("table", {"class": "TeacherStatisticsTable"})
    t_body = table.find("tbody")

    trs = t_body.find_all("tr")
    data = get_teacher_data(trs)
    write_data_to_file(data, f"{data['Teacher']}_{date_to.strftime('%B')},{year}", date_to)
