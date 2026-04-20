# data_processing页面
import numbers
import pandas as pd
from pandas.api.types import is_numeric_dtype


def _is_numeric_like(value):
    if pd.isna(value):
        return False

    if isinstance(value, numbers.Number):
        return True

    text = str(value).strip()
    if text == "":
        return False

    try:
        float(text)
        return True
    except (TypeError, ValueError):
        return False


def validate_excel(excel_data):
    if "住院号" not in excel_data.columns:
        return False

    series = excel_data["住院号"]

    if series.isna().any():
        return False

    if is_numeric_dtype(series):
        return True

    return series.map(_is_numeric_like).all()