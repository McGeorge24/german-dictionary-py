from UrediCelice import *


def NovPridevnik(tabela, cursor: CellCoord) -> None:
    pridevnik = str(input("Vnesi pridevnik:\t")).lower()
    pomen = str(input("Vnesi pomen/prevod:\t")).lower()

    # v primeru da uporabnik proces prekine
    if pridevnik == "cancel" or pomen == "cancel":
        return None

    # najde prvo prosto vrstico
    vrstica = NajdiCelicoStolpec(tabela, cursor)

    # shrani vse podatke o novem pridevniku
    tabela[vrstica.format()] = pridevnik
    vrstica.col += 1
    tabela[vrstica.format()] = pomen


def NovGlagol(tabela, cursor: CellCoord) -> None:
    glagol = str(input("Vnesi glagol:\t")).lower()
    third_person = str(input("Vnesi 3. osebo ednine:\t")).lower()
    perfekt = str(input("Vnesi obliko v perfekt\t")).lower()
    pomen = str(input("Vnesi pomen/prevod:\t")).lower()

    # v primeru da uporabnik proces prekine
    if glagol == "cancel" or third_person == "cancel" or perfekt == "cancel" or pomen == "cancel":
        return None

    # najde prvo prosto vrstico
    vrstica = NajdiCelicoStolpec(tabela, cursor)

    # shrani vse podatke o novem glagolu
    tabela[vrstica.format()] = glagol
    vrstica.col += 1
    tabela[vrstica.format()] = third_person
    vrstica.col += 1
    tabela[vrstica.format()] = perfekt
    vrstica.col += 1
    tabela[vrstica.format()] = pomen


def NovPrislov(tabela, cursor: CellCoord) -> None:
    prislov = str(input("Vnesi prislov:\t")).lower()
    pomen = str(input("Vnesi pomen/prevod:\t")).lower()

    # v primeru da uporabnik proces prekine
    if prislov == "cancel" or pomen == "cancel":
        return None

    # najde prvo prosto vrstico
    vrstica = NajdiCelicoStolpec(tabela, cursor)

    # shrani nov prislov
    tabela[vrstica.format()] = prislov
    vrstica.col += 1
    tabela[vrstica.format()] = pomen


def NovaBesednaZveza(tabela, cursor: CellCoord) -> None:
    nova_zveza = str(input("Vnesi novo besedno zvezo:\t"))
    pomen = str(input("Vnesi pomen/prevod:\t")).lower()

    # v primeru da uporabnik proces prekine
    if nova_zveza.lower() == "cancel" or pomen == "cancel":
        return None

    # najde prvo prosto vrstico
    vrstica = NajdiCelicoStolpec(tabela, cursor)

    # shrani novo besedno zvezo in pomen
    tabela[vrstica.format()] = nova_zveza
    vrstica.col += 1
    tabela[vrstica.format()] = pomen
