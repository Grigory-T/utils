def header_info(df):
    """Return one row per input column as a pandas or eager Polars DataFrame.

    The output uses the input backend. idx is zero-based, and col name is the
    input label as text. row_n is the number of input rows. notna count and
    isna count follow the flat-table blank rules. unique count uses export
    precision and counts pandas types separately.
    Row 1–3 and the first five unique examples are text previews. Flags are 1/0.
    Requires pandas or Polars for the corresponding input.
    """
    return _analyze_columns(df, detail=False)


def check_dtypes(df):
    """Return column stats and one row per nonblank elements_dtype per column.

    Accepts and returns a pandas or eager Polars DataFrame. row_n is the number
    of input rows. Counts repeat for mixed columns. memory, MB is estimated
    in decimal MB once per input column.
    str key stats gives mi, mx, me, st (min, max, mean, and sample SD of length).
    Low/high and shortest/longest give up to three unique examples per type.
    all_unique(unique_justseen(els)) checks for repeated runs within str or int.
    Char count is exact. Characters is capped at Excel's cell limit.
    Flags are 1/0. An all-blank column has one row with blank type-specific stats.
    """
    return _analyze_columns(df, detail=True)


def _analyze_columns(df, *, detail):
    from datetime import date, datetime
    from heapq import nlargest, nsmallest
    from math import isnan
    import re

    try:
        import polars as pl
    except ImportError:
        pl = None
    polars_frame = pl is not None and isinstance(df, pl.DataFrame)
    pd = None
    if not polars_frame:
        try:
            import pandas as pd
        except ImportError:
            raise TypeError("Expected a pandas or Polars DataFrame") from None
        if not isinstance(df, pd.DataFrame):
            raise TypeError("Expected a pandas or Polars DataFrame")
        import numpy as np

    def scalar(value):
        if pd is not None:
            if value is pd.NA or value is pd.NaT:
                return None
            if isinstance(value, np.datetime64):
                value = None if np.isnat(value) else pd.Timestamp(value)
            elif isinstance(value, np.timedelta64):
                raise TypeError("Timedelta values are unsupported")
            elif isinstance(value, np.generic):
                value = value.item()
            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime(warn=False)
            if isinstance(value, float) and isnan(value):
                return None
        if isinstance(value, str):
            return value.replace("\r\n", "\n").replace("\r", "\n") or None
        if isinstance(value, datetime):
            if value.utcoffset() is not None:
                raise ValueError("Timezone-aware datetimes are unsupported")
            return value.replace(microsecond=value.microsecond // 1000 * 1000)
        return value

    def preview(value):
        if value is None:
            return None
        if isinstance(value, str):
            return repr(value[:55]) + ("…" if len(value) > 55 else "")
        if isinstance(value, datetime):
            precision = "milliseconds" if value.microsecond else "seconds"
            return value.isoformat(sep=" ", timespec=precision)
        return str(value)

    def examples(values):
        return " · ".join(map(preview, values)) or None

    def polars_text(series):
        values = series.cast(pl.String)
        if values.str.contains("\r", literal=True).any():
            values = values.str.replace_all(r"\r\n?", "\n")
        return values.filter(values != "")

    def groups_for(series):
        dtype = series.dtype
        if polars_frame:
            base = dtype.base_type()
            if base in (pl.String, pl.Categorical, pl.Enum):
                return {"str": polars_text(series)}, None
            values = series.drop_nulls()
            if base == pl.Null:
                return {}, None
            if base == pl.Datetime and not dtype.time_zone:
                return {"datetime": values.dt.truncate("1ms")}, None
            if dtype.is_integer():
                return {"int": values}, None
            if base in (pl.Float32, pl.Float64):
                if values.is_nan().any():
                    raise ValueError("Polars NaN is unsupported")
                return {"float": values}, None
            if base in (pl.Boolean, pl.Date):
                return {"bool" if base == pl.Boolean else "date": values}, None
            raise TypeError(f"Unsupported dtype: {dtype}")

        values = series.dropna()
        kinds = {"i": "int", "u": "int", "f": "float", "b": "bool", "M": "datetime"}
        if isinstance(dtype, pd.DatetimeTZDtype) or dtype.kind in ("m", "c"):
            raise TypeError(f"Unsupported dtype: {dtype}")
        if values.empty:
            return {}, None
        if dtype.kind in kinds and not isinstance(dtype, pd.CategoricalDtype):
            kind = "date" if dtype.type is date else kinds[dtype.kind]
            if kind == "datetime":
                if values.dt.tz is not None:
                    raise TypeError(f"Unsupported dtype: {dtype}")
                values = values.dt.floor("ms")
            return {kind: values}, None
        if pd.api.types.infer_dtype(values.astype(object), skipna=False) == "string":
            if pl is not None:
                strings = pl.Series(values.to_numpy(dtype=object), dtype=pl.String)
                return {"str": polars_text(strings)}, None
            return {"str": [scalar(v) for v in values if v != ""]}, None

        groups = {}
        first = {}
        for value in values:
            value = scalar(value)
            if value is None:
                continue
            if not isinstance(value, (str, bool, int, float, date)):
                raise TypeError(f"Unsupported cell value: {type(value).__name__}")
            kind = type(value).__name__
            groups.setdefault(kind, []).append(value)
            if not detail and len(first) < 5:
                first[(kind, value)] = value
        return groups, list(first.values())

    def text_metrics(values, unique, native):
        if native:
            lengths = values.str.len_chars()
            unique_lengths = unique.str.len_chars()
            shortest = unique.gather(unique_lengths.arg_sort().head(3))
            longest = unique.gather(unique_lengths.arg_sort(descending=True).head(3))
            sd = lengths.std(ddof=1) or 0.0
        else:
            lengths = np.fromiter(map(len, values), dtype=np.int64, count=len(values))
            shortest = nsmallest(3, unique, key=len)
            longest = nlargest(3, unique, key=len)
            sd = lengths.std(ddof=1) if len(values) > 1 else 0.0
        maximum = int(lengths.max())
        chars = set()
        unseen = None
        step = max(1, min(4096, 2**20 // maximum))
        for start in range(0, len(unique), step):
            batch = unique[start:start + step]
            text = batch.str.join("").item() if native else "".join(batch)
            before = len(chars)
            chars.update(text if unseen is None else unseen.findall(text))
            # Scan only unseen characters when the alphabet is small.
            if len(chars) != before:
                unseen = re.compile("[^" + re.escape("".join(chars)) + "]") if len(chars) <= 128 else None
        alphabet = repr("".join(sorted(chars)))
        encoded = alphabet.encode("utf-16-le")
        if len(encoded) > 65534:
            alphabet = encoded[:65532].decode("utf-16-le", "ignore") + "…"
        str_key_stats = f"mi={int(lengths.min())} mx={maximum} me={lengths.mean():.15g} st={sd:.15g}"
        return str_key_stats, examples(shortest), examples(longest), len(chars), alphabet

    def describe(kind, values):
        native = pl is not None and isinstance(values, pl.Series)
        pandas_series = pd is not None and isinstance(values, pd.Series)
        if native:
            unique = values.unique(maintain_order=True)
        elif pandas_series:
            unique = pd.Series(values.unique())
        else:
            unique = list(dict.fromkeys(values))
        count, unique_count = len(values), len(unique)
        if not detail:
            return count, unique_count, list(unique[:5])

        low = high = true = false = None
        if kind == "bool":
            true = int(values.sum()) if native or pandas_series else sum(values)
            false = count - true
        elif native:
            low = examples(unique.bottom_k(3).sort())
            high = examples(unique.top_k(3).sort(descending=True))
        elif pandas_series:
            if kind in ("date", "datetime") and isinstance(unique.dtype, getattr(pd, "ArrowDtype", ())):
                ordered = unique.sort_values()
                low, high = examples(ordered.iloc[:3]), examples(ordered.iloc[-3:][::-1])
            else:
                low, high = examples(unique.nsmallest(3)), examples(unique.nlargest(3))
        else:
            low, high = examples(nsmallest(3, unique)), examples(nlargest(3, unique))

        all_unique_justseen = None
        if kind in ("str", "int"):
            if unique_count in (1, count):
                all_unique_justseen = True
            elif native:
                all_unique_justseen = values.rle_id().max() + 1 == unique_count
            elif pandas_series:
                array = values.to_numpy()
                all_unique_justseen = (array[1:] != array[:-1]).sum() + 1 == unique_count
            else:
                all_unique_justseen = 1 + sum(a != b for a, b in zip(values, values[1:])) == unique_count
        str_key_stats, *text = text_metrics(values, unique, native) if kind == "str" else (None,) * 5
        flag = None if all_unique_justseen is None else int(all_unique_justseen)
        properties = (kind, count, unique_count, int(count == unique_count), flag, low, high, true, false, *text)
        return count, unique_count, (str_key_stats, *properties)

    columns = {
        "idx": int, "col name": str, "row_n": int,
        "unique count": int, "notna count": int, "isna count": int,
    }
    if detail:
        elements_columns = {
            "elements_dtype": str, "elements count": int, "elements unique count": int,
            "elements all_unique_flag": int, "all_unique(unique_justseen(els))": int,
            "Low 3": str, "High 3": str, "True count": int, "False count": int,
            "Shortest 3": str, "Longest 3": str,
            "Char count": int, "Characters": str,
        }
        columns.update({
            "memory, MB": float, "str key stats": str, "series_dtype": str,
            **elements_columns, "all_unique_flag": int,
        })
    else:
        columns.update({"dtype": str, "examples": str, "all_unique_flag": int,
                        "Row 1": str, "Row 2": str, "Row 3": str})

    row_n = df.shape[0]
    rows = []
    for idx, col_name in enumerate(df.columns):
        series = df.to_series(idx) if polars_frame else df.iloc[:, idx]
        groups, first = groups_for(series)
        stats = [describe(kind, values) for kind, values in groups.items() if len(values)]
        notna_count = sum(count for count, _, _ in stats)
        unique_count = sum(count for _, count, _ in stats)
        counts = (idx, str(col_name), row_n, unique_count, notna_count, row_n - notna_count)
        series_dtype = str(series.dtype)
        all_unique_flag = int(row_n > 0 and unique_count == row_n)
        if detail:
            memory_bytes = (
                series.estimated_size() if polars_frame else series.memory_usage(index=False, deep=True)
            )
            types = [properties for _, _, properties in stats] or [(None,) * (1 + len(elements_columns))]
            for i, (str_key_stats, *elements) in enumerate(types):
                memory_mb = memory_bytes / 10**6 if i == 0 else None
                rows.append((*counts, memory_mb, str_key_stats, series_dtype, *elements, all_unique_flag))
        else:
            sample = [preview(scalar(value)) for value in series.head(3).to_list()]
            sample.extend([None] * (3 - len(sample)))
            unique_examples = first if first is not None else stats[0][2] if stats else []
            rows.append((*counts, series_dtype, examples(unique_examples), all_unique_flag, *sample))

    if polars_frame:
        dtypes = {int: pl.Int64, str: pl.String, float: pl.Float64}
        schema = {name: dtypes[typ] for name, typ in columns.items()}
        return pl.DataFrame(rows, schema=schema, orient="row")
    dtypes = {int: "Int64", str: "string", float: "Float64"}
    return pd.DataFrame(rows, columns=list(columns)).astype({name: dtypes[typ] for name, typ in columns.items()})
