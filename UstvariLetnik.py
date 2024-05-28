import openpyxl as excel
import openpyxl.styles as styles


class Style:
    def __init__(self, tip: str):
        self.tip = tip
        # za naslov
        if tip.lower() == "naslov":
            self.font = styles.Font(
                name='Arial', size=20, bold=True, color='000000')
            self.fill = styles.PatternFill(
                start_color='AA88FF', end_color='AA88FF', fill_type='solid')
            self.border = styles.Border(
                left=styles.Side("thick"), right=styles.Side("thick"), top=styles.Side("thick"), bottom=styles.Side("thick"))
            self.alignment = styles.Alignment(
                horizontal='center', vertical='center')
        # za vse ostalo
        else:
            self.font = styles.Font(name='Arial', size=12, color='000000')
            self.fill = None
            self.border = None
            self.alignment = styles.Alignment(
                horizontal='left', vertical='center')


def NaslovPridevnik(tabela, cursor_pos: str, style: Style) -> None:

    tabela["B3"] = "PRIDEVNIKI"
    tabela.merge_cells('B3:C3')
    print(style.tip)
    tabela["B3"].font = style.font
    tabela["B3"].fill = style.fill
    tabela["B3"].border = style.border
    tabela["C3"].border = style.border
    tabela["B3"].alignment = style.alignment


def NaslovGlagol() -> None:
    print()


def NaslovPrislov() -> None:
    print()


def NaslovDrugo() -> None:
    print()


def NovLetnik(workbook, trenuten_letnik: int) -> None:
    workbook.create_sheet(title=f"{trenuten_letnik+1}.letnik")
    nov_letnik = workbook[f"{trenuten_letnik+1}.letnik"]

    slog_naslov = Style("naslov")
    nov_letnik["A1"] = "Letnik:"
    nov_letnik["B1"] = trenuten_letnik+1
    cursor_pos = "B3"

    NaslovPridevnik(nov_letnik, cursor_pos, slog_naslov)
    NaslovGlagol()
    NaslovPrislov()
    NaslovDrugo()
