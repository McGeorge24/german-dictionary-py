from UstvariLetnik import *
from DodajanjeBesed import *


class Program:

    def __init__(self) -> None:
        # prebere trenuten letnik iz datoteke
        with open('assets\\trenutno_stanje.txt', 'r') as file:
            # .strip() removes the newline
            self.trenuten_letnik = int(file.readline().strip())

        # generira kazalec
        self.kazalec = CellCoord(3, 2)
        # odpre zvezek
        self.zvezek = excel.load_workbook('assets\\test.xlsx')
        # odpre list
        if self.trenuten_letnik != 0:
            self.tabela = self.zvezek[f"{self.trenuten_letnik}.letnik"]
        # če letnik = 0 (ga še ni) naredi novega
        else:
            NovLetnik(self.zvezek, self.trenuten_letnik)
            self.trenuten_letnik += 1
            self.tabela = self.zvezek[f"{self.trenuten_letnik}.letnik"]

    # shrani zvezek (deluje kot destructor)
    def save(self):
        self.zvezek.save('assets\\test.xlsx')
        with open('assets\\trenutno_stanje.txt', 'w') as file:
            file.write(f"{self.trenuten_letnik}")

    # glavna zanka
    def run(self):
        while True:
            # resetira kazalec na polje pridevnik (prvo polje vseh tabel)
            self.kazalec.row = 3
            self.kazalec.col = 2

            # sprejme ukaz
            temp_input = str(
                input("Vnesi ukaz (new tip_besede ali exit):\t")).lower()

            # izvede po navodilih
            if temp_input == "exit" or temp_input == "save":
                break

            elif "prid" in temp_input:
                self.kazalec = NajdiCelicoVrsta(
                    self.tabela, self.kazalec, vrednost="PRIDEVNIKI")
                NovPridevnik(self.tabela, self.kazalec)

            elif temp_input == "glagol":
                self.kazalec = NajdiCelicoVrsta(
                    self.tabela, self.kazalec, vrednost="GLAGOLI")
                NovGlagol(self.tabela, self.kazalec)

            elif temp_input == "prislov":
                self.kazalec = NajdiCelicoVrsta(
                    self.tabela, self.kazalec, vrednost="PRISLOVI")
                NovPrislov(self.tabela, self.kazalec)

            elif temp_input == "drugo":
                self.kazalec = NajdiCelicoVrsta(
                    self.tabela, self.kazalec, vrednost="DRUGE OBLIKE")
                NovaBesednaZveza(self.tabela, self.kazalec)

            elif temp_input == "tema":
                DodajTemo(self.tabela, self.kazalec)

            elif "samo" in temp_input:
                NovSamostalnik(self.tabela, self.kazalec)

            elif temp_input == "letnik":
                NovLetnik(self.zvezek, self.trenuten_letnik)
                self.trenuten_letnik += 1
