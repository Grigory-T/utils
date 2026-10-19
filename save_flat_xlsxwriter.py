def save_flat_xlsxwriter(df, path, title="") -> None:
    """Save a Polars DataFrame to a new/overwritten .xlsx flat table; return None.

    Dependencies: polars, XlsxWriter. Example: save_flat_xlsxwriter(df, "out.xlsx").
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
    Streams worksheet rows with constant_memory; the input df stays in memory.
    """
    import math
    import re
    from datetime import date, datetime
    from pathlib import Path

    import polars as pl
    import xlsxwriter

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

    headers = ["№пп", *df.columns, "last"]
    last_col = len(headers)
    invalid_text = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")
    for value in [f"Таблица: {title}", *headers]:
        if invalid_text.search(value) or len(value.encode("utf-16-le")) > 65534:
            raise ValueError("Text exceeds Excel limits or contains XML controls")

    with xlsxwriter.Workbook(
        str(path),
        {
            "constant_memory": True,
            "strings_to_formulas": False,
            "strings_to_urls": False,
            "strings_to_numbers": False,
        },
    ) as wb:
        ws = wb.add_worksheet("Data")
        ws.hide_gridlines(2)
        ws.set_zoom(85)
        ws.set_tab_color("#002060")
        ws.set_default_row(14.5)
        ws.set_paper(9)
        ws.set_portrait()
        ws.set_margins(0.7, 0.7, 0.75, 0.75)
        ws.set_column(0, last_col, 5.0)
        if df.width:
            ws.set_column(2, last_col - 1, 11.83)

        base = {"font_name": "Calibri", "font_size": 11}
        body = wb.add_format({**base, "border": 1})
        center = wb.add_format(
            {**base, "border": 1, "align": "center", "valign": "vcenter"}
        )
        heading = {
            **base,
            "border": 1,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        }
        blue = wb.add_format({**heading, "bg_color": "#00B0F0"})
        black = wb.add_format(
            {**heading, "bg_color": "#000000", "font_color": "#FFFFFF"}
        )
        bold = wb.add_format({**base, "bold": True})
        date_fmt = wb.add_format({**base, "border": 1, "num_format": "yyyy-mm-dd"})
        time_fmt = wb.add_format(
            {**base, "border": 1, "num_format": "yyyy-mm-dd hh:mm:ss"}
        )

        if "_x" in title:
            parts = [s for s in re.split("(_)", f"Таблица: {title}") if s]
            ws.write_rich_string(1, 1, *[v for s in parts for v in (bold, s)], bold)
        else:
            ws.write_string(1, 1, f"Таблица: {title}", bold)
        for col in range(1, last_col + 1):
            ws.write_formula(3, col, "=COLUMN()-COLUMN($A$4)", center, col)
        ws.set_row(5, 48.65)
        for col, name in enumerate(headers, 1):
            header_fmt = black if col in (1, last_col) else blue
            if "_x" in name:
                parts = [s for s in re.split("(_)", name) if s]
                ws.write_rich_string(
                    5, col, *[v for s in parts for v in (header_fmt, s)], header_fmt
                )
            else:
                ws.write_string(5, col, name, header_fmt)

        for row, values in enumerate(df.iter_rows(), 6):
            ws.write_formula(row, 1, "=ROW()-ROW($B$6)", center, row - 5)
            for col, value in enumerate(values, 2):
                if isinstance(value, str):
                    if (
                        invalid_text.search(value)
                        or len(value.encode("utf-16-le")) > 65534
                    ):
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
                fmt = (
                    time_fmt
                    if isinstance(value, datetime)
                    else date_fmt
                    if isinstance(value, date)
                    else body
                )
                if isinstance(value, str) and "_x" in value:
                    # Split underscores to preserve literal OOXML escape sequences.
                    parts = [s for s in re.split("(_)", value) if s]
                    ws.write_rich_string(
                        row, col, *[v for s in parts for v in (fmt, s)], fmt
                    )
                else:
                    ws.write(row, col, value, fmt)
            ws.write_number(row, last_col, 1, center)

        ws.autofilter(5, 1, max(5, df.height + 5), last_col)
