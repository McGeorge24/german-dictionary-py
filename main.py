import openpyxl as excel
from UstvariLetnik import *
from interface import Program

trenuten_letnik = 1


def main() -> None:
    program = Program(trenuten_letnik)
    program.run()
    program.save()


if __name__ == "__main__":
    main()
