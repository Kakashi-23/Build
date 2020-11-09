import datetime
import schedule
from xlrd import open_workbook
from xlrd import xldate_as_tuple
from pyautogui import hotkey
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from twilio.rest import Client
import time

file_path = r"time.xlsx"
file = open_workbook(file_path)
sheet_main = file.sheet_by_index(0)
#  Code to allow access for Microphone, Camera and notifications
# 0 is disable and 1 is allow.
opt = Options()
opt.add_argument("--disable-infobars")
opt.add_argument("start-maximized")
opt.add_argument("--disable-extensions")
# Pass the argument 1 to allow and 2 to block
opt.add_experimental_option("prefs", {
    "profile.default_content_setting_values.media_stream_mic": 1,
    "profile.default_content_setting_values.media_stream_camera": 1,
    "profile.default_content_setting_values.geolocation": 1,
    "profile.default_content_setting_values.notifications": 1
})


def notify_class_ended():
    account_sid = 'AC27aab58fe36b9113e1d341848427c806'
    auth_token = '284a82b189ce68ea30c089960d53894a'
    client = Client(account_sid, auth_token)

    client.messages.create(
        from_='whatsapp:+14155238886',
        body='Your class had ended',
        to='whatsapp:+919532709622'
    )


def class_joined():
    account_sid = 'AC27aab58fe36b9113e1d341848427c806'
    auth_token = '284a82b189ce68ea30c089960d53894a'
    client = Client(account_sid, auth_token)

    client.messages.create(
        from_='whatsapp:+14155238886',
        body='class joined successfully',
        to='whatsapp:+919532709622'
    )


def start_meet(meet_id, driver):
    driver.get(meet_id)
    time.sleep(5)
    hotkey('ctrl', 'e')
    time.sleep(5)
    hotkey('ctrl', 'd')
    time.sleep(60)
    try:
        join = driver.find_element_by_xpath(
            "//*[@id='yDmH0d']/c-wiz/div/div/div[6]/div[3]/div/div/div[2]/div/div[1]/div[2]/div/div[2]/div/div[1]/div[1]/span/span")

        time.sleep(300)
        join.click()
        class_joined()
    except NoSuchElementException:
        driver.close()
        notify_no_class("joining time error")
        pass
    time.sleep(3000)
    driver.close()
    notify_class_ended()


def get_meet_id(subject_fun):
    subject_fun = str(subject_fun).title()
    meet_id = ''
    file_path1 = r"Google Meet IDs.xlsx"
    file1 = open_workbook(file_path1)
    sheet = file1.sheet_by_index(0)
    for i in range(1, sheet.nrows):
        if subject_fun in sheet.cell_value(i, 1):
            meet_id = sheet.cell_value(i, 2)
            break
    login(meet_id)


def notify_no_class(reason):
    account_sid = 'AC27aab58fe36b9113e1d341848427c806'
    auth_token = '284a82b189ce68ea30c089960d53894a'
    client = Client(account_sid, auth_token)

    client.messages.create(
        from_='whatsapp:+14155238886',
        body=str(reason),
        to='whatsapp:+919532709622'
    )


def start_class(sheet, day, class_time):
    try:
        subject = sheet.cell_value(find_row(sheet, day), find_column(sheet, class_time))
        if len(str(subject)) == 0:
            notify_no_class("no class at this hour")
        else:
            get_meet_id(subject)
    except TypeError:

        notify_no_class('Error block')


def find_column(sheet, class_time):
    for i in range(1, sheet.ncols):
        time1 = xldate_as_tuple(sheet_main.cell_value(0, i), file.datemode)
        class_time1 = str(datetime.time(*time1[3:]))
        if class_time == class_time1:
            return i


def find_row(sheet, day):
    for i in range(1, sheet.nrows):
        if day in sheet.cell_value(i, 0):
            return i


def scheduling(sheet1, file1):
    day = datetime.datetime.now().strftime("%A")
    for i in range(1, sheet1.ncols):
        if day.lower() == "monday":
            time1 = xldate_as_tuple(sheet1.cell_value(0, i), file1.datemode)
            class_time = str(datetime.time(*time1[3:]))
            schedule.every().monday.at(class_time).do(start_class, sheet1, day, class_time)
        if day.lower() == "tuesday":
            time1 = xldate_as_tuple(sheet1.cell_value(0, i), file1.datemode)
            class_time = str(datetime.time(*time1[3:]))
            schedule.every().tuesday.at(class_time).do(start_class, sheet1, day, class_time)
        if day.lower() == "wednesday":
            time1 = xldate_as_tuple(sheet1.cell_value(0, i), file1.datemode)
            class_time = str(datetime.time(*time1[3:]))
            schedule.every().wednesday.at(class_time).do(start_class, sheet1, day, class_time)
        if day.lower() == "thursday":
            time1 = xldate_as_tuple(sheet1.cell_value(0, i), file1.datemode)
            class_time = str(datetime.time(*time1[3:]))
            schedule.every().thursday.at(class_time).do(start_class, sheet1, day, class_time)
        if day.lower() == "friday":
            time1 = xldate_as_tuple(sheet1.cell_value(0, i), file1.datemode)
            class_time = str(datetime.time(*time1[3:]))
            schedule.every().friday.at(class_time).do(start_class, sheet1, day, class_time)
        if day.lower() == "saturday":
            notify_no_class()
        if day.lower() == "sunday":
            notify_no_class()


def login(meet_id):
    driver = webdriver.Chrome(options=opt,
                              executable_path=r"C:\Users\Ideapad\Desktop\python\chromedriver_win32\chromedriver.exe")
    driver.get(
        "https://accounts.google.com/ServiceLogin?service=mail&passive=true&rm=false&continue=https://mail.google.com/mail/&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1#identifier")
    time.sleep(4)
    driver.find_element_by_xpath("//input[@name='identifier']").send_keys("17372@nith.ac.in")
    time.sleep(2)
    # Next Button:
    driver.find_element_by_xpath("//*[@id='identifierNext']/div/button/div[2]").click()
    time.sleep(5)
    # Password:
    driver.find_element_by_xpath("//input[@name='password']").send_keys("kakashi@234")
    time.sleep(2)
    # next button:
    driver.find_element_by_xpath("//*[@id='passwordNext']/div/button").click()
    time.sleep(5)
    start_meet(meet_id, driver)


def main_callback():
    scheduling(sheet_main, file)
    print(schedule.jobs)
    while True:
        schedule.run_pending()
        time.sleep(5)

if __name__ == '__main__':
    main_callback()