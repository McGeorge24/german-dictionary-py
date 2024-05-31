import openpyxl as excel
import openpyxl.styles as styles
from UrediCelice import *
from UstvariLetnik import *
from DodajanjeBesed import *


class Program:
    def __init__(self, trenuten_letnik: int) -> None:
        self.trenuten_letnik = trenuten_letnik
        self.cursor = CellCoord(1, 1)
        self.zvezek = excel.load_workbook('assets\\test.xlsx')
        if self.trenuten_letnik != 0:
            self.tabela = self.zvezek[f"{trenuten_letnik}.letnik"]
        else:
            NovLetnik(self.zvezek, self.trenuten_letnik)
            trenuten_letnik += 1
            self.tabela = self.zvezek[f"{trenuten_letnik}.letnik"]

    def save(self):
        self.zvezek.save('assets\\test.xlsx')

    def run(self):
        while True:
            # resetira kazalec na polje pridevnik (prvo polje vseh tabel)
            self.cursor.row = 3
            self.cursor.col = 2

            # sprejme ukaz
            temp_input = str(
                input("Vnesi ukaz (new tip_besede ali exit):\t")).lower()

            # izvede po navodilih
            if temp_input == "exit":
                break
            elif temp_input == "new pridevnik":
                self.cursor = NajdiCelicoVrsta(
                    self.tabela, self.cursor, vrednost="PRIDEVNIKI")
                NovPridevnik(self.tabela, self.cursor)
            elif temp_input == "new glagol":
                self.cursor = NajdiCelicoVrsta(
                    self.tabela, self.cursor, vrednost="GLAGOLI")
                NovGlagol(self.tabela, self.cursor)
            elif temp_input == "new prislov":
                self.cursor = NajdiCelicoVrsta(
                    self.tabela, self.cursor, vrednost="PRISLOVI")
                NovPrislov(self.tabela, self.cursor)
            elif temp_input == "new drugo":
                self.cursor = NajdiCelicoVrsta(
                    self.tabela, self.cursor, vrednost="DRUGE OBLIKE")
                NovaBesednaZveza(self.tabela, self.cursor)
