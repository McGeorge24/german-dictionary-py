from openpyxl import load_workbook

def Naslov() -> None:
    print()

def NovLetnik() -> None:
    print()

def NovLetnik() -> None:
    print()

def main() -> None:
    # Naloži zvezek in list
    workbook = load_workbook('assets\\test_spreadsheet.xlsx')
    letnik = workbook["LETNIK1"]

    # Manipuliranje s celicami
    letnik["A1"] = "Hello excel"
    print("hello world")

    # Shrani zvezek
    workbook.save('assets\\test_spreadsheet.xlsx')


if __name__ == "__main__":
    main()
