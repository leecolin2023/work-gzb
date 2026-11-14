"""Index the existing bank extracts and collateral worksheets by real coordinates."""
from pathlib import Path

from openpyxl import load_workbook


def read_workbook(path, *, source=None):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        sheets = {}
        for sheet in book:
            rows = []
            for row in sheet.iter_rows():
                values = {cell.column_letter: cell.value for cell in row if cell.value is not None}
                if values:
                    rows.append(dict(row=next(c.row for c in row if c.value is not None), cells=values))
            sheets[sheet.title] = rows
        return dict(source=source or Path(path).name, sheets=sheets,
                    quality="已存整理表/摘录。空值不等于零，权利与余额以相应原件进一步核验。")
    finally:
        book.close()
