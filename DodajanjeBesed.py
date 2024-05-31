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


# pridobi seznam vseh tem (za samostalnike)
def PridobiTeme(tabela, zacetna_celica: CellCoord) -> list:
    cursor = zacetna_celica
    meja = cursor.col+100
    seznam_tem = []
    while cursor != CellCoord(1, 2):
        cursor = NajdiCelicoVrsta(
            tabela, cursor, vrednost="samostalniki", absolutna_meja=meja)
        cursor.col += 1
        print(f"{cursor.row}, {cursor.col}")
        seznam_tem.append(tabela[cursor.format()].value)
    seznam_tem.pop()
    return (seznam_tem)


def DodajTemo(tabela, cursor: CellCoord) -> None:
    cursor = NajdiCelicoVrsta(tabela, cursor, vrednost="samostalniki")
    naslov = Style("naslov")
    podnaslov = Style("podnaslov")
    # najdi prosto mesto če še ni bila dodana nobena tema
    print(cursor == CellCoord(1, 1))
    if cursor == CellCoord(1, 1):
        cursor.row = 3
        cursor = NajdiCelicoVrsta(tabela, cursor, vrednost="DRUGE OBLIKE")
        cursor.col += 3
    # najdi prvo temo in prištevaj 5 (toliko prostora porabi 1 tema), dokler ne najdeš prostega mesta
    else:
        while tabela[cursor.format()].value != None:
            cursor.col += 5
    tabela[cursor.format()] = "samostalniki"
    cursor.col += 1
    tabela[cursor.format()] = str(input("Vnesi novo temo:\t"))
    tabela.merge_cells(CellRange(cursor, CellCoord(cursor.row, cursor.col+2)))
    cursor.col -= 1
    ApplyStyleToCells(tabela, cursor.row, cursor.col,
                      cursor.row, cursor.col+3, naslov)
    tabela[cursor.format()].font = styles.Font(
        name='Arial', size=12, bold=False, color='000000')

    cursor.row += 1
    tabela[cursor.format()] = "člen"
    cursor.col += 1
    tabela[cursor.format()] = "samostalnik"
    cursor.col += 1
    tabela[cursor.format()] = "množina"
    cursor.col += 1
    tabela[cursor.format()] = "pomen"
    ApplyStyleToCells(tabela, cursor.row, cursor.col - 3,
                      cursor.row, cursor.col, podnaslov)


def NovSamostalnik(tabela, cursor: CellCoord) -> None:
    seznam_tem = PridobiTeme(tabela, cursor)
    for i, tema in enumerate(seznam_tem):
        print(f"{i}) {tema}")
    izbrana_tema = seznam_tem[int(
        input("Vnesi številko teme, kateri hočeš dodati samostalnik;\t"))]
    cursor.col = 1
    cursor = NajdiCelicoVrsta(tabela, cursor, vrednost=izbrana_tema)
    cursor.col -= 1

    cursor = NajdiCelicoStolpec(tabela, cursor)
    clen = str(input("Vnesi clen samostalnika:\t")).lower()
    samostalnik = str(input("Vnesi samostalnik:\t")).capitalize()
    mnozina = "-".join(str(input("Vnesi mnozinsko koncnico\t-")))
    pomen = str(input("Vnesi pomen/prevod:\t")).lower()
    tabela[cursor.format()] = clen
    cursor.col += 1
    tabela[cursor.format()] = samostalnik
    cursor.col += 1
    tabela[cursor.format()] = mnozina
    cursor.col += 1
    tabela[cursor.format()] = pomen
    cursor.col += 1
