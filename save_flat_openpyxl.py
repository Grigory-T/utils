def save_flat_openpyxl(df, path, title="") -> None:
    """Save a Polars DataFrame to a new/overwritten .xlsx flat table; return None.

    Dependencies: polars, openpyxl>=3.1. Example: save_flat_openpyxl(df, "out.xlsx").
    path: str or PathLike; parent must exist. title: str.
    Columns: String/Categorical/Enum, Boolean, integer, Float32/64, Date,
    timezone-naive Datetime, Null. Dates must be >= 1900-03-01.
    No LazyFrame, Series, nested or Object data.
    Null/empty strings become blank; CRLF/CR become LF in data strings.
    Datetimes truncate to milliseconds.
    Require finite floats in Excel's normal range, integers within
    +/-999999999999999, and XML-safe
    strings up to 32767 UTF-16 units. Other Excel numeric precision limits apply.
    Invalid values raise; a failed write may leave a partial output file.
    """
    import math
    import re
    from datetime import date, datetime
    from pathlib import Path

    import polars as pl
    from openpyxl import Workbook
    from openpyxl.cell.rich_text import CellRichText
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    if not isinstance(df, pl.DataFrame):
        raise TypeError("df must be a Polars DataFrame")
    path = Path(path)
    if path.suffix.lower() != ".xlsx" or not isinstance(title, str):
        raise ValueError("Use an .xlsx path and a string title")
    if df.height > 1048570 or df.width > 16381:
        raise ValueError("Data exceeds Excel sheet limits")
    if {"№пп", "last"} & set(df.columns):
        raise ValueError("Reserved columns: №пп, last")
    for dtype in df.dtypes:
        allowed = dtype.is_integer() or dtype.base_type() in (
            pl.String,
            pl.Categorical,
            pl.Enum,
            pl.Boolean,
            pl.Float32,
            pl.Float64,
            pl.Date,
            pl.Datetime,
            pl.Null,
        )
        if not allowed or (dtype.base_type() == pl.Datetime and dtype.time_zone):
            raise TypeError(f"Unsupported dtype: {dtype}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = ws.sheet_view.zoomScaleNormal = 85
    ws.sheet_properties.tabColor = "002060"
    ws.sheet_format.defaultRowHeight = 14.5
    ws.row_dimensions[6].height = 48.65
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_margins.left = ws.page_margins.right = 0.7
    ws.page_margins.top = ws.page_margins.bottom = 0.75
    ws.page_margins.header = ws.page_margins.footer = 0.3
    ws.column_dimensions["A"].width = 5.81640625

    font = Font(name="Calibri", size=11)
    bold = Font(name="Calibri", size=11, bold=True)
    side = Side(style="thin")
    border = Border(left=side, right=side, top=side, bottom=side)
    center = Alignment(horizontal="center", vertical="center")
    wrapped = Alignment(horizontal="center", vertical="center", wrap_text=True)
    headers = ["№пп", *df.columns, "last"]
    last_col = len(headers) + 1
    invalid_text = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")

    for value in [f"Таблица: {title}", *headers]:
        if invalid_text.search(value) or len(value.encode("utf-16-le")) > 65534:
            raise ValueError("Text exceeds Excel limits or contains XML controls")
    ws["B2"] = CellRichText([s for s in re.split("(_)", f"Таблица: {title}") if s])
    ws["B2"].font = bold
    for col, name in enumerate(headers, 2):
        boundary = col in (2, last_col)
        ws.column_dimensions[get_column_letter(col)].width = (
            5.81640625 if boundary else 12.54296875
        )
        cell = ws.cell(4, col, "=COLUMN()-COLUMN($A$4)")
        cell.font, cell.border, cell.alignment = font, border, center
        cell = ws.cell(6, col, CellRichText([s for s in re.split("(_)", name) if s]))
        cell.data_type = "s"
        cell.font = Font(
            name="Calibri", size=11, bold=True, color="FFFFFF" if boundary else "000000"
        )
        cell.fill = PatternFill("solid", fgColor="000000" if boundary else "00B0F0")
        cell.border, cell.alignment = border, wrapped

    for row, values in enumerate(df.iter_rows(), 7):
        for col, value in enumerate([None, *values, 1], 2):
            if isinstance(value, str):
                if invalid_text.search(value) or len(value.encode("utf-16-le")) > 65534:
                    raise ValueError(
                        "Text exceeds Excel limits or contains XML controls"
                    )
                value = value.replace("\r\n", "\n").replace("\r", "\n")
            elif isinstance(value, float) and (
                not math.isfinite(value)
                or abs(value) > 9.99999999999999e307
                or 0 < abs(value) < 2.2250738585072014e-308
            ):
                raise ValueError("Float is outside Excel's finite normal range")
            elif isinstance(value, int) and abs(value) > 999999999999999:
                raise ValueError("Integer exceeds Excel's 15-digit limit")
            if isinstance(value, datetime):
                value = value.replace(microsecond=value.microsecond // 1000 * 1000)
            if isinstance(value, date) and value < (
                datetime(1900, 3, 1)
                if isinstance(value, datetime)
                else date(1900, 3, 1)
            ):
                raise ValueError("Dates before 1900-03-01 are unsupported")
            if isinstance(value, str) and "_x" in value:
                # Split underscores to preserve literal OOXML escape sequences.
                value = CellRichText([s for s in re.split("(_)", value) if s])
            cell = ws.cell(row, col, value)
            if isinstance(value, (str, CellRichText)):
                cell.data_type = "s"
            cell.font, cell.border = font, border
            if isinstance(value, date):
                cell.number_format = (
                    "yyyy-mm-dd hh:mm:ss"
                    if isinstance(value, datetime)
                    else "yyyy-mm-dd"
                )
            if col in (2, last_col):
                cell.alignment = center
        ws.cell(row, 2, "=ROW()-ROW($B$6)")

    ws.auto_filter.ref = f"B6:{get_column_letter(last_col)}{max(6, df.height + 6)}"
    wb.save(path)
