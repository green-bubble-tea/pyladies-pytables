from openpyxl import Workbook, load_workbook
import os

FILENAME = "workbook1.xlsx"


def read_workbook(filename):
    workbook: Workbook = load_workbook(os.path.join("workshop", filename))
    try:
        return workbook
    finally:
        workbook.close()


def set_name(sheet, index):
    sheet.title = f"sheet_{index + 1}"
    print(sheet.title)


def write_value(book):
    for index, sheet in enumerate(book.worksheets):
        sheet[f"A{index + 1}"] = f"Test_{index + 1}"


if __name__ == "__main__":
    w = read_workbook(FILENAME)
    try:
        [set_name(i, index) for index, i in enumerate(w)]
        write_value(w)
        sheet = w["sheet_1"]
        for row in sheet:
            [print(cell.value, end="\t") for cell in row]
            print("\n")
    finally:
        w.save(os.path.join("workshop", FILENAME))
