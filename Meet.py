from webbrowser import get
from xlrd import open_workbook
import PyPDF2


def start_meet(meet_id):
    path = r'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
    driver = get(path)
    driver.open(meet_id)


def get_meet_id(subject_fun):
    subject_fun = str(subject_fun).title()
    meet_id = ''
    file_path = r"Google Meet IDs.xlsx"
    file = open_workbook(file_path)
    sheet = file.sheet_by_index(0)
    for i in range(1, sheet.nrows):
        if subject_fun in sheet.cell_value(i, 1):
            meet_id = sheet.cell_value(i, 2)
            break

    start_meet(meet_id)



