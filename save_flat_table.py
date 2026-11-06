def save_flat_table(tables, path=None, title="") -> None:
    """Write pandas/Polars tables to a new/overwritten .xlsx file; return None.

    Dependencies: XlsxWriter and pandas and/or polars; neither frame library
    requires the other. path: str/PathLike; parent must exist. None saves
    flat_table.xlsx beside the running script, or in cwd in Jupyter/interactive use.
    tables: DataFrame, nonempty list/tuple of DataFrames, or {sheet_name: df}.
    A frame uses Sheet1; a sequence uses Sheet1, Sheet2, ...; dict order is kept.
    title: str for all sheets, list/tuple in sheet order, or {sheet_name: str}.
    Sheet names must be valid, unique ignoring case, and at most 31 UTF-16 units.
    Escape-like names (_xNNNN_) are rejected because Excel silently renames them.

    Columns: String/Categorical/Enum, Boolean, integer, Float32/64, Date,
    timezone-naive Datetime, Null. Dates must be >= 1900-03-01.
    pandas object columns accept these scalar types; its index is ignored.
    Headers become unique strings without mutating frames. No Series, LazyFrame,
    nested values, arbitrary objects, Decimal, Time or Duration data.
    Null/empty strings and pandas NaN/NA/NaT become blank; Polars NaN is rejected.
    Data-string CRLF/CR becomes LF.
    Datetimes truncate to milliseconds. Require Excel-range finite floats,
    integers within +/-999999999999999, and XML-safe text <=32767 UTF-16 units.
    Other Excel numeric precision limits apply. Invalid writes may leave a
    partial file. Rows stream with constant_memory; input frames stay in memory.
    Every sheet has width 5 from A through 20 columns beyond the table (up to
    XFD), freezes rows 1-6 at A7, and filters the complete table starting at B6.

    save_flat_table(df)
    save_flat_table({"Orders": orders, "Lines": lines}, "out.xlsx",
                    title={"Orders": "Orders", "Lines": "Order lines"})
    """
    import math
    import re
    import sys
    from datetime import date, datetime
    from pathlib import Path

    import xlsxwriter

    try:
        import polars as pl
    except ImportError:
        pl = None
    try:
        import pandas as pd
    except ImportError:
        pd = None
    frame_types = tuple(lib.DataFrame for lib in (pl, pd) if lib is not None)
    if not frame_types:
        raise ImportError("Install pandas or polars to supply DataFrames")

    if isinstance(tables, frame_types):
        sheets = {"Sheet1": tables}
    elif isinstance(tables, (list, tuple)):
        sheets = {f"Sheet{i}": df for i, df in enumerate(tables, 1)}
    elif isinstance(tables, dict):
        sheets = dict(tables)
    else:
        raise TypeError("tables must be a pandas/Polars DataFrame, list/tuple or dict")
    if not sheets:
        raise ValueError("At least one table is required")
    if isinstance(title, str):
        titles = dict.fromkeys(sheets, title)
    elif isinstance(title, (list, tuple)):
        if len(title) != len(sheets):
            raise ValueError("Provide one title per sheet")
        titles = dict(zip(sheets, title))
    elif isinstance(title, dict):
        if title.keys() != sheets.keys():
            raise ValueError("Title keys must match sheet names")
        titles = dict(title)
    else:
        raise TypeError("title must be a string, list/tuple or dict")

    if path is None:
        script = getattr(sys.modules.get("__main__"), "__file__", None)
        folder = (
            Path(script).resolve().parent
            if script and not script.startswith("<")
            else Path.cwd()
        )
        path = folder / "flat_table.xlsx"
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Use an .xlsx output path")
    invalid_text = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")
    seen = set()
    sheet_headers = {}
    for name, df in sheets.items():
        if not isinstance(name, str):
            raise TypeError("Sheet names must be strings")
        if (
            not name.strip()
            or len(name.encode("utf-16-le")) > 62
            or re.search(r"[\[\]:*?/\\]", name)
            or re.search(r"_x[0-9A-Fa-f]{4}_", name)
            or invalid_text.search(name)
            or name.startswith("'")
            or name.endswith("'")
            or name.casefold() == "history"
            or name.casefold() in seen
        ):
            raise ValueError(f"Invalid or duplicate Excel sheet name: {name!r}")
        seen.add(name.casefold())
        if not isinstance(df, frame_types):
            raise TypeError(f"Sheet {name!r} must contain a pandas/Polars DataFrame")
        if not isinstance(titles[name], str):
            raise TypeError("Titles must be strings")
        height, width = df.shape
        if height > 1048570 or width > 16381:
            raise ValueError(f"Sheet {name!r} exceeds Excel limits")
        headers, used = [], set()
        for label in df.columns:
            base = header = str(label)
            suffix = 1
            while header in used:
                header = f"{base}_{suffix}"
                suffix += 1
            headers.append(header)
            used.add(header)
        sheet_headers[name] = headers
        if {"№пп", "last"} & set(headers):
            raise ValueError("Reserved columns: №пп, last")
        if pl is not None and isinstance(df, pl.DataFrame):
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
                if not allowed or (
                    dtype.base_type() == pl.Datetime and dtype.time_zone
                ):
                    raise TypeError(f"Unsupported dtype: {dtype}")
        else:
            for dtype in df.dtypes:
                if (
                    isinstance(dtype, pd.DatetimeTZDtype)
                    or pd.api.types.is_timedelta64_dtype(dtype)
                    or pd.api.types.is_complex_dtype(dtype)
                ):
                    raise TypeError(f"Unsupported dtype: {dtype}")
        for value in [name, f"Таблица: {titles[name]}", *headers]:
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

        for name, df in sheets.items():
            sheet_title = titles[name]
            headers = ["№пп", *sheet_headers[name], "last"]
            last_col = len(headers)
            pandas_frame = pd is not None and isinstance(df, pd.DataFrame)
            if pandas_frame:
                import numpy as np
            ws = wb.add_worksheet(name)
            ws.hide_gridlines(2)
            ws.set_zoom(85)
            ws.set_tab_color("#002060")
            ws.set_default_row(14.5)
            ws.set_paper(9)
            ws.set_portrait()
            ws.set_margins(0.7, 0.7, 0.75, 0.75)
            ws.set_column(0, min(16383, last_col + 20), 5)
            ws.freeze_panes(6, 0)

            if "_x" in sheet_title:
                parts = [s for s in re.split("(_)", f"Таблица: {sheet_title}") if s]
                ws.write_rich_string(1, 1, *[v for s in parts for v in (bold, s)], bold)
            else:
                ws.write_string(1, 1, f"Таблица: {sheet_title}", bold)
            for col in range(1, last_col + 1):
                ws.write_formula(3, col, "=COLUMN()-COLUMN($A$4)", center, col)
            ws.set_row(5, 48.65)
            for col, header in enumerate(headers, 1):
                header_fmt = black if col in (1, last_col) else blue
                if "_x" in header:
                    parts = [s for s in re.split("(_)", header) if s]
                    ws.write_rich_string(
                        5, col, *[v for s in parts for v in (header_fmt, s)], header_fmt
                    )
                else:
                    ws.write_string(5, col, header, header_fmt)

            if df.shape[1] == 0:
                rows = (() for _ in range(df.shape[0]))
            else:
                rows = (
                    df.itertuples(index=False, name=None)
                    if pandas_frame
                    else df.iter_rows()
                )
            for row, values in enumerate(rows, 6):
                ws.write_formula(row, 1, "=ROW()-ROW($B$6)", center, row - 5)
                for col, value in enumerate(values, 2):
                    if pandas_frame:
                        if value is pd.NA or value is pd.NaT:
                            value = None
                        elif isinstance(value, np.datetime64):
                            value = (
                                None
                                if np.isnat(value)
                                else pd.Timestamp(value).to_pydatetime(warn=False)
                            )
                        elif isinstance(value, np.timedelta64):
                            raise TypeError("Timedelta values are unsupported")
                        elif isinstance(value, np.generic):
                            value = value.item()
                        if isinstance(value, float) and math.isnan(value):
                            value = None
                        if isinstance(value, pd.Timestamp):
                            value = value.to_pydatetime(warn=False)
                    if value is not None and not isinstance(
                        value, (str, bool, int, float, date)
                    ):
                        raise TypeError(
                            f"Unsupported cell value: {type(value).__name__}"
                        )
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
                        if value.tzinfo is not None and value.utcoffset() is not None:
                            raise ValueError("Timezone-aware datetimes are unsupported")
                        value = value.replace(
                            microsecond=value.microsecond // 1000 * 1000
                        )
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

            ws.autofilter(5, 1, max(5, df.shape[0] + 5), last_col)
