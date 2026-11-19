def save_flat_table(tables, path=None, title="", header_analysis=False) -> None:
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
    Hard values only: row/column numbers are literal integers, never formulas.
    Formula-like input strings stay text. Every column A:XFD has width 5.
    Data sheets use 85% zoom, freeze rows 1-6 at A7, and filter from B6.
    header_analysis: bool, False by default. True adds header and header_dtype
    views for each table, prefixed by the data sheet name for multiple tables.
    They include three source rows, distinct examples, counts, value types,
    column memory in MiB, and text length, character and range summaries.
    Counts and examples exclude exported blanks; distinctness is per value type.

    save_flat_table(df)
    save_flat_table({"Orders": orders, "Lines": lines}, "out.xlsx",
                    title={"Orders": "Orders", "Lines": "Order lines"})
    """
    import math
    import re
    import sys
    from shutil import copyfileobj
    from tempfile import TemporaryDirectory
    from zipfile import ZipFile
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
            ws.set_column(0, 16383, 5)
            ws.freeze_panes(6, 0)

            if "_x" in sheet_title:
                parts = [s for s in re.split("(_)", f"Таблица: {sheet_title}") if s]
                ws.write_rich_string(1, 1, *[v for s in parts for v in (bold, s)], bold)
            else:
                ws.write_string(1, 1, f"Таблица: {sheet_title}", bold)
            for col in range(1, last_col + 1):
                ws.write_number(3, col, col, center)
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
                ws.write_number(row, 1, row - 5, center)
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

        if header_analysis:
            from heapq import nlargest, nsmallest
            from textwrap import wrap as wrap_text

            def sample_value(value, pandas_column=False):
                if pandas_column:
                    if value is pd.NA or value is pd.NaT:
                        return None
                    if isinstance(value, np.datetime64):
                        if np.isnat(value):
                            return None
                        value = pd.Timestamp(value)
                    elif isinstance(value, np.generic):
                        value = value.item()
                    if isinstance(value, pd.Timestamp):
                        value = value.to_pydatetime(warn=False)
                if isinstance(value, str):
                    return value.replace("\r\n", "\n").replace("\r", "\n") or None
                if isinstance(value, float) and math.isnan(value):
                    return None
                if isinstance(value, datetime):
                    return value.replace(microsecond=value.microsecond // 1000 * 1000)
                return value

            def show(value, limit=24):
                if isinstance(value, str):
                    return repr(value[:limit]) + ("…" if len(value) > limit else "")
                if isinstance(value, datetime):
                    precision = "milliseconds" if value.microsecond else "seconds"
                    return value.isoformat(sep=" ", timespec=precision)
                return str(value)

            def ends(label1, low, label2, high):
                return "\n".join(
                    f"{label} " + " · ".join(show(v, 12) for v in values)
                    for label, values in ((label1, low), (label2, high)) if len(values)
                )

            def describe(values, kind):
                native = pl is not None and isinstance(values, pl.Series)
                pandas_series = pd is not None and isinstance(values, pd.Series)
                count = len(values)
                unique = (
                    values.unique(maintain_order=True)
                    if native
                    else pd.Series(values.unique())
                    if pandas_series
                    else list(dict.fromkeys(values))
                )
                distinct = len(unique)
                examples = list(unique[:5])
                types = f"{kind}\n{count:,} / {distinct:,}"
                lengths = extremes = lengths_extremes = characters = ""
                if kind == "bool":
                    true = int(values.sum()) if native or pandas_series else sum(values)
                    extremes = f"True {true:,}\nFalse {count - true:,}"
                elif native:
                    extremes = ends(
                        "Low", unique.bottom_k(3).sort(),
                        "High", unique.top_k(3).sort(descending=True),
                    )
                elif pandas_series:
                    arrow_dtype = getattr(pd, "ArrowDtype", ())
                    if kind in ("date", "datetime") and isinstance(unique.dtype, arrow_dtype):
                        ordered = unique.sort_values()
                        extremes = ends("Low", ordered.iloc[:3], "High", ordered.iloc[-3:][::-1])
                    else:
                        extremes = ends("Low", unique.nsmallest(3), "High", unique.nlargest(3))
                else:
                    extremes = ends("Low", nsmallest(3, unique), "High", nlargest(3, unique))

                if kind == "str":
                    if native:
                        length = values.str.len_chars()
                        unique_length = unique.str.len_chars()
                        shortest = unique.gather(unique_length.arg_sort().head(3))
                        longest = unique.gather(unique_length.arg_sort(descending=True).head(3))
                        deviation = length.std(ddof=1) or 0.0
                        minimum, maximum, mean = length.min(), length.max(), length.mean()
                    else:
                        length = np.fromiter(map(len, values), dtype=np.int64, count=count)
                        shortest = nsmallest(3, unique, key=len)
                        longest = nlargest(3, unique, key=len)
                        deviation = length.std(ddof=1) if count > 1 else 0.0
                        minimum, maximum, mean = length.min(), length.max(), length.mean()
                    lengths = (
                        f"Min {minimum:,} / max {maximum:,}\n"
                        f"Mean {mean:,.1f} / SD {deviation:,.1f}"
                    )
                    lengths_extremes = ends("Short", shortest, "Long", longest)
                    chars = set()
                    unseen = None
                    step = min(4096, 2**20 // int(maximum))
                    for start in range(0, distinct, step):
                        chunk = unique[start:start + step]
                        text = chunk.str.join("").item() if native else "".join(chunk)
                        previous = len(chars)
                        chars.update(text if unseen is None else unseen.findall(text))
                        # Skip already counted characters in columns with small alphabets.
                        if previous != len(chars):
                            unseen = (
                                re.compile("[^" + re.escape("".join(chars)) + "]")
                                if len(chars) <= 128 else None
                            )
                    visible = "".join(nsmallest(60, chars))
                    characters = f"{len(chars):,}\n{visible!r}" + ("…" if len(chars) > 60 else "")
                columns = [types, lengths, extremes, lengths_extremes, characters]
                return count, distinct, examples, columns

            def polars_strings(values):
                if values.str.contains("\r", literal=True).any():
                    values = values.str.replace_all(r"\r\n?", "\n")
                values = values.filter(values != "")
                return [describe(values, "str")] if len(values) else []

            def profile(series):
                if pl is not None and isinstance(series, pl.Series):
                    dtype = series.dtype
                    memory = series.estimated_size("mb")
                    if dtype.base_type() in (pl.String, pl.Categorical, pl.Enum):
                        groups = polars_strings(series.cast(pl.String))
                    else:
                        values = series.drop_nulls()
                        if dtype.base_type() == pl.Datetime:
                            values = values.dt.truncate("1ms")
                            kind = "datetime"
                        elif dtype.is_integer():
                            kind = "int"
                        else:
                            kind = {pl.Boolean: "bool", pl.Date: "date"}.get(dtype, "float")
                        groups = [describe(values, kind)] if len(values) else []
                else:
                    memory = series.memory_usage(index=False, deep=True) / (1024 ** 2)
                    values = series.dropna()
                    dtype = series.dtype
                    kinds = {"i": "int", "u": "int", "f": "float", "b": "bool", "M": "datetime"}
                    if values.empty:
                        groups = []
                    elif dtype.kind in kinds and not isinstance(dtype, pd.CategoricalDtype):
                        kind = "date" if dtype.type is date else kinds[dtype.kind]
                        if kind == "datetime":
                            values = values.dt.floor("ms")
                        groups = [describe(values, kind)]
                    elif pd.api.types.infer_dtype(values.astype(object), skipna=False) == "string":
                        if pl is not None:
                            strings = pl.Series(values.to_numpy(dtype=object), dtype=pl.String)
                            groups = polars_strings(strings)
                        else:
                            strings = [sample_value(v) for v in values if v != ""]
                            groups = [describe(strings, "str")] if strings else []
                    else:
                        typed = {}
                        first = {}
                        for value in values:
                            value = sample_value(value, pandas_column=True)
                            if value is not None:
                                kind = type(value).__name__
                                typed.setdefault(kind, []).append(value)
                                if len(first) < 5:
                                    first[(kind, value)] = value
                        groups = [describe(group, kind) for kind, group in typed.items()]
                        return str(dtype), memory, list(first.values()), groups
                return str(dtype), memory, groups[0][2] if groups else [], groups

            def analysis_name(name, label):
                prefix = f"{name} " if len(sheets) > 1 else ""
                encoded = prefix.encode("utf-16-le")
                suffix = label
                number = 2
                while True:
                    available = 31 - len(suffix)
                    trimmed = encoded[:available * 2].decode("utf-16-le", "ignore")
                    candidate = (trimmed + suffix).lstrip("'")
                    if candidate.casefold() not in seen:
                        seen.add(candidate.casefold())
                        return candidate
                    suffix = f"{label}_{number}"
                    number += 1

            wrap = wb.add_format({**base, "border": 1, "text_wrap": True, "valign": "vcenter"})
            numeric = {**base, "border": 1, "valign": "vcenter"}
            count_fmt = wb.add_format({**numeric, "num_format": "#,##0"})
            memory_fmt = wb.add_format({**numeric, "num_format": "0.00#"})
            sample_date = wb.add_format({**numeric, "num_format": "yyyy-mm-dd"})
            sample_time = wb.add_format({**numeric, "num_format": "yyyy-mm-dd hh:mm:ss.000"})

            def row_height(cells, widths):
                lines = max(
                    sum(max(1, len(wrap_text(line, width - 2))) for line in cell.split("\n"))
                    for cell, width in zip(cells, widths)
                )
                return max(22, min(106, 5 + 14 * lines))

            def write_cell(ws, row, col, value, fmt):
                if isinstance(value, str) and "_x" in value:
                    parts = [s for s in re.split("(_)", value) if s]
                    ws.write_rich_string(row, col, *[v for s in parts for v in (fmt, s)], fmt)
                else:
                    ws.write(row, col, value, fmt)

            for name, df in sheets.items():
                rows, columns = df.shape
                pandas_column = pd is not None and isinstance(df, pd.DataFrame)
                summary = wb.add_worksheet(analysis_name(name, "header"))
                detail = wb.add_worksheet(analysis_name(name, "header_dtype"))
                sample_widths = (20, 23, 23, 23)
                summary_widths = (16, 9, 9, 9, 9, 34)
                detail_widths = (8, 17, 18, 26, 25, 20)
                summary_labels = ("Dtype", "Distinct", "Present", "Blank", "All unique", "Distinct examples")
                detail_labels = (
                    "MiB", "Type\nPresent / distinct", "Text length", "Low / high",
                    "Shortest / longest", "Characters",
                )
                for ws, labels, widths in (
                    (summary, summary_labels, summary_widths),
                    (detail, detail_labels, detail_widths),
                ):
                    ws.hide_gridlines(2)
                    ws.set_zoom(85)
                    ws.freeze_panes(1, 3)
                    ws.set_row(0, 44)
                    for col, width in enumerate((4, 1, *sample_widths, 1, *widths)):
                        ws.set_column(col, col, width)
                    ws.write_string(0, 0, "#", blue)
                    write_cell(ws, 0, 2, f"{name}\n{rows:,} rows · {columns + 2:,} fields", blue)
                    for col in range(3, 6):
                        ws.write_string(0, col, f"Row {col - 2}", blue)
                    ws.write_row(0, 7, labels, blue)

                detail_row = 1
                for index, field in enumerate(["№пп", *sheet_headers[name], "last"], 1):
                    technical = index in (1, columns + 2)
                    if technical:
                        numbered = index == 1
                        distinct = rows if numbered else int(rows > 0)
                        samples = [i if numbered else 1 for i in range(1, min(rows, 3) + 1)]
                        examples = list(range(1, min(rows, 5) + 1)) if numbered else samples[:1]
                        high = list(range(rows, max(0, rows - 3), -1)) if numbered else examples
                        low = samples if numbered else examples
                        dtype, memory, present = "int (generated)", None, rows
                        properties = []
                        if rows:
                            extremes = ends("Low", low, "High", high)
                            properties.append([f"int\n{rows:,} / {distinct:,}", "", extremes, "", ""])
                    else:
                        series = df.iloc[:, index - 2] if pandas_column else df.to_series(index - 2)
                        dtype, memory, examples, groups = profile(series)
                        present = sum(group[0] for group in groups)
                        distinct = sum(group[1] for group in groups)
                        samples = [sample_value(v, pandas_column) for v in series.head(3).to_list()]
                        properties = [group[3] for group in groups]

                    example_text = " · ".join(show(v) for v in examples)
                    sample_text = [field[:72], *(v[:66] for v in samples if isinstance(v, str))]
                    sample_height = row_height(sample_text, sample_widths)
                    summary_height = row_height((dtype, example_text), (summary_widths[0], summary_widths[-1]))
                    summary.set_row(index, max(sample_height, summary_height))

                    for ws, row in ((summary, index), (detail, detail_row)):
                        ws.write_number(row, 0, index, center)
                        write_cell(ws, row, 2, field, black if technical else blue)
                        for col, value in enumerate((samples + [None] * 3)[:3], 3):
                            fmt = (
                                sample_time if isinstance(value, datetime)
                                else sample_date if isinstance(value, date)
                                else wrap
                            )
                            write_cell(ws, row, col, value, fmt)

                    summary.write_string(index, 7, dtype, wrap)
                    summary.write_row(index, 8, (distinct, present, rows - present), count_fmt)
                    summary.write(index, 11, "Yes" if rows and distinct == rows else None, center)
                    write_cell(summary, index, 12, example_text, wrap)
                    detail.write(detail_row, 7, memory, memory_fmt)
                    for offset, cells in enumerate(properties or [["", "", "", "", ""]]):
                        row = detail_row + offset
                        detail.set_row(row, max(sample_height, row_height(cells, detail_widths[1:])))
                        if offset:
                            detail.write_number(row, 0, index, center)
                            write_cell(detail, row, 2, field, blue)
                        for col, value in enumerate(cells, 8):
                            write_cell(detail, row, col, value, wrap)
                    detail_row += max(1, len(properties))

    # OOXML includes padding: 5.81640625 stores Excel width 5 in Calibri 11.
    # Stream-copy the workbook, changing column metadata only.
    with TemporaryDirectory(dir=path.parent) as folder:
        fixed = Path(folder) / "flat_table.xlsx"
        with ZipFile(path) as source, ZipFile(fixed, "w") as target:
            for info in source.infolist():
                with source.open(info) as reader, target.open(info, "w") as writer:
                    sheet = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", info.filename)
                    if sheet and (not header_analysis or int(sheet[1]) <= len(sheets)):
                        header, marker, data = reader.read(65536).partition(
                            b"<sheetData>"
                        )
                        header, count = re.subn(
                            rb'(<col\b[^>]*\bwidth=")[^"]+(")',
                            rb"\g<1>5.81640625\2",
                            header,
                        )
                        if not marker or count != 1:
                            raise ValueError("Unexpected worksheet column layout")
                        writer.write(header + marker + data)
                    copyfileobj(reader, writer, 1024 * 1024)
        fixed.replace(path)
