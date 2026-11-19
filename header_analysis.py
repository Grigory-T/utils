def analyze_header(df):
    """Return one row per input field as a pandas or eager Polars DataFrame.

    The output uses the input backend. Position is 1-based and Field is the
    input label as text. Counts follow flat-table blanks, newlines and datetime
    precision. Distinctness keeps types separate in mixed pandas columns.
    Row 1–3 and the first five distinct examples are text previews. Flags are 1/0.
    Requires pandas or Polars for the corresponding input.
    """
    return _analyze_columns(df, detail=False)


def analyze_header_dtype(df):
    """Return column stats and one row per present value type in each field.

    Accepts and returns a pandas or eager Polars DataFrame. Column counts
    repeat for mixed fields; estimated memory (MiB) appears once per field.
    Low/high and shortest/longest contain up to three distinct examples.
    One run checks whether each distinct string/integer occurs in a single
    consecutive run within the present values of its type.
    Text length SD is sample SD (zero for a single value). Text character
    counts are exact; the sorted inventory is capped at Excel's cell limit.
    Flags are 1/0. An all-blank field has one row with blank type-specific metrics.
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
        return (
            int(lengths.min()), maximum, float(lengths.mean()), float(sd),
            examples(shortest), examples(longest), len(chars), alphabet,
        )

    def describe(kind, values):
        native = pl is not None and isinstance(values, pl.Series)
        pandas_series = pd is not None and isinstance(values, pd.Series)
        if native:
            unique = values.unique(maintain_order=True)
        elif pandas_series:
            unique = pd.Series(values.unique())
        else:
            unique = list(dict.fromkeys(values))
        count, distinct = len(values), len(unique)
        if not detail:
            return count, distinct, list(unique[:5])

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

        one_run = None
        if kind in ("str", "int"):
            if distinct in (1, count):
                one_run = True
            elif native:
                one_run = values.rle_id().max() + 1 == distinct
            elif pandas_series:
                array = values.to_numpy()
                one_run = bool((array[1:] != array[:-1]).sum() + 1 == distinct)
            else:
                one_run = 1 + sum(a != b for a, b in zip(values, values[1:])) == distinct
        text = text_metrics(values, unique, native) if kind == "str" else (None,) * 8
        flag = None if one_run is None else int(one_run)
        return count, distinct, (kind, count, distinct, int(count == distinct), flag, low, high, true, false, *text)

    columns = {
        "Position": int, "Field": str, "Dtype": str,
        "Distinct": int, "Present": int, "Blank": int, "All unique": int,
    }
    if detail:
        type_columns = {
            "Value type": str, "Type count": int, "Type distinct": int,
            "Type uniq?": int, "One run": int, "Low 3": str, "High 3": str,
            "True count": int, "False count": int, "Min length": int, "Max length": int,
            "Mean length": float, "Length SD": float, "Shortest 3": str, "Longest 3": str,
            "Char count": int, "Characters": str,
        }
        columns.update({"MiB": float, **type_columns})
    else:
        columns.update({"Row 1": str, "Row 2": str, "Row 3": str, "Examples": str})

    rows = []
    for position, label in enumerate(df.columns, 1):
        series = df.to_series(position - 1) if polars_frame else df.iloc[:, position - 1]
        groups, first = groups_for(series)
        stats = [describe(kind, values) for kind, values in groups.items() if len(values)]
        present = sum(count for count, _, _ in stats)
        distinct = sum(count for _, count, _ in stats)
        height = len(series)
        field = (
            position, str(label), str(series.dtype), distinct, present,
            height - present, int(height > 0 and distinct == height),
        )
        if detail:
            mib = series.estimated_size("mb") if polars_frame else series.memory_usage(index=False, deep=True) / 2**20
            types = [properties for _, _, properties in stats] or [(None,) * len(type_columns)]
            rows.extend((*field, mib if i == 0 else None, *properties) for i, properties in enumerate(types))
        else:
            sample = [preview(scalar(value)) for value in series.head(3).to_list()]
            sample.extend([None] * (3 - len(sample)))
            distinct_examples = first if first is not None else stats[0][2] if stats else []
            rows.append((*field, *sample, examples(distinct_examples)))

    if polars_frame:
        dtypes = {int: pl.Int64, str: pl.String, float: pl.Float64}
        schema = {name: dtypes[typ] for name, typ in columns.items()}
        return pl.DataFrame(rows, schema=schema, orient="row")
    dtypes = {int: "Int64", str: "string", float: "Float64"}
    return pd.DataFrame(rows, columns=list(columns)).astype({name: dtypes[typ] for name, typ in columns.items()})
