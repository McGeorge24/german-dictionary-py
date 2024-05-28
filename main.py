import openpyxl as excel
from UstvariLetnik import *

trenuten_letnik = 0


def main() -> None:
    # Naloži zvezek in list
    workbook = excel.load_workbook('assets\\test.xlsx')
    letnik = workbook["Sheet1"]

    # Manipuliranje s celicami
    letnik["A1"] = "Hello excel"

    NovLetnik(workbook, trenuten_letnik)

    # Shrani zvezek
    workbook.save('assets\\test.xlsx')


if __name__ == "__main__":
    main()
