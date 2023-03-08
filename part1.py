from openpyxl import load_workbook
import codecs
from dateutil.parser import parse
import csv
import os


def person_to_string(t, d):
    name = ' '.join(t[:2])
    dep = d[t[2]]
    return f"{name}, {dep}, {', '.join(map(str, t[3:]))}"


def exercise_1():
    dir_name = "lesson_01"
    workers = load_workbook(os.path.join(dir_name, "Seznam pracovníků - aktualizováno k 14.11.2019.xlsx")).worksheets[0]
    with codecs.open(os.path.join(dir_name, "seznam_oddeleni.csv")) as dep_file:
        department = csv.DictReader(dep_file, delimiter=";")
        departments = {int(i["Cislo oddeleni"]): i["Nazev oddeleni"] for i in department}
    persons = {}
    for index, value in enumerate(workers.values):
        if index > 0:
            person_id = int(value[0])
            persons[person_id] = value[1:]
    with codecs.open(os.path.join(dir_name, "log_pristupu_trezor.csv")) as log_file:
        log = csv.DictReader(log_file, delimiter=";")
        for i in log:
            date_ = parse(i["cas_pristupu"])
            person_id = int(i["osobni cislo"])
            who_did = persons[person_id]
            dep = who_did[2]
            if (date_.hour < 6 or date_.hour > 21) and departments[dep] != "Reditel":
                print(f"{person_id}, "
                      f"{person_to_string(who_did, departments)}, "
                      f"{i['datum_pristupu']}, "
                      f"{i['cas_pristupu']}")


if __name__ == "__main__":
    exercise_1()
