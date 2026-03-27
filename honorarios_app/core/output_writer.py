from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
BOLD = Font(bold=True)


def style_header(ws):
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(1, c)
        cell.fill = HEADER_FILL
        cell.font = BOLD


def autofit_columns(ws, max_width=60):
    for col in range(1, ws.max_column + 1):
        max_len = 0
        for row in range(1, min(ws.max_row, 2000) + 1):
            v = ws.cell(row, col).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(col)].width = min(max(10, max_len + 2), max_width)


def add_totals_row(ws, start_row, idx_saldo, idx_vlr_aut, label_col=1):
    last_data_row = ws.max_row
    totals_row = last_data_row + 2

    ws.cell(totals_row, label_col).value = "TOTALES"
    ws.cell(totals_row, label_col).font = BOLD

    if idx_saldo:
        c = get_column_letter(idx_saldo)
        ws.cell(totals_row, idx_saldo).value = f"=SUM({c}{start_row}:{c}{last_data_row})"
        ws.cell(totals_row, idx_saldo).font = BOLD

    if idx_vlr_aut:
        c = get_column_letter(idx_vlr_aut)
        ws.cell(totals_row, idx_vlr_aut).value = f"=SUM({c}{start_row}:{c}{last_data_row})"
        ws.cell(totals_row, idx_vlr_aut).font = BOLD