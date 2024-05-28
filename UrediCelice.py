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
            thick = styles.Side("thick")
            self.border = styles.Border(
                left=thick, right=thick, top=thick, bottom=thick)
            self.alignment = styles.Alignment(
                horizontal='center', vertical='center')
        # za vpodnaslov
        elif tip.lower() == "podnaslov":
            self.font = styles.Font(name='Arial', size=12, color='000000')
            self.fill = styles.PatternFill()
            thin = styles.Side("thin")
            self.border = styles.Border(left=thin, right=thin, bottom=thin)
            self.alignment = styles.Alignment(
                horizontal="center", vertical="center")
        # za vse ostalo
        else:
            self.font = styles.Font(name='Calibri', size=12, color='000000')
            self.fill = None
            self.border = None
            self.alignment = styles.Alignment(
                horizontal='left', vertical='center')


def ApplyStyleToCells(sheet, start_row, start_col, end_row, end_col, style: Style) -> None:
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell = sheet.cell(row=row, column=col)
            cell.font = style.font
            cell.fill = style.fill
            cell.border = style.border
            cell.alignment = style.alignment
