import openpyxl as excel
import openpyxl.styles as styles
from UrediCelice import *


def NaslovPridevnik(tabela, cursor_pos: tuple, naslov: Style, podnaslov: Style) -> None:
    tabela["B3"] = "PRIDEVNIKI"
    tabela.merge_cells('B3:C3')
    ApplyStyleToCells(tabela, 3, 2, 3, 3, naslov)
    tabela["B4"] = "pridevnik"
    tabela["C4"] = "pomen"
    ApplyStyleToCells(tabela, 4, 2, 4, 3, podnaslov)


def NaslovGlagol(tabela, cursor_pos: tuple, naslov: Style, podnaslov: Style) -> None:
    tabela["E3"] = "GLAGOLI"
    tabela.merge_cells('E3:H3')
    ApplyStyleToCells(tabela, 3, 5, 3, 8, naslov)
    tabela["E4"] = "glagol"
    tabela["F4"] = "3. oseba"
    tabela["G4"] = "perfekt?"
    tabela["H4"] = "pomen"
    ApplyStyleToCells(tabela, 4, 5, 4, 8, podnaslov)


def NaslovPrislov(tabela, cursor_pos: tuple, naslov: Style, podnaslov: Style) -> None:
    tabela["J3"] = "PRISLOVI"
    tabela.merge_cells('J3:K3')
    ApplyStyleToCells(tabela, 3, 10, 3, 11, naslov)
    tabela["J4"] = "prislov"
    tabela["K4"] = "pomen"
    ApplyStyleToCells(tabela, 4, 10, 4, 11, podnaslov)


def NaslovDrugo(tabela, cursor_pos: tuple, naslov: Style, podnaslov: Style) -> None:
    tabela["M3"] = "DRUGE OBLIKE"
    tabela.merge_cells('M3:N3')
    ApplyStyleToCells(tabela, 3, 13, 3, 14, naslov)
    tabela["M4"] = "besedna zveza"
    tabela["N4"] = "pomen"
    ApplyStyleToCells(tabela, 4, 13, 4, 14, podnaslov)


def NovLetnik(workbook, trenuten_letnik: int) -> None:
    workbook.create_sheet(title=f"{trenuten_letnik+1}.letnik")
    nov_letnik = workbook[f"{trenuten_letnik+1}.letnik"]

    slog_naslov = Style("naslov")
    slog_podnaslov = Style("podnaslov")
    nov_letnik["A1"] = "Letnik:"
    nov_letnik["B1"] = trenuten_letnik+1
    cursor_pos = (3, 2)

    NaslovPridevnik(nov_letnik, cursor_pos, slog_naslov, slog_podnaslov)
    NaslovGlagol(nov_letnik, cursor_pos, slog_naslov, slog_podnaslov)
    NaslovPrislov(nov_letnik, cursor_pos, slog_naslov, slog_podnaslov)
    NaslovDrugo(nov_letnik, cursor_pos, slog_naslov, slog_podnaslov)
