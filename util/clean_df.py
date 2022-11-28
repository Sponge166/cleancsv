import pandas as pd

def is_nan_series(col: pd.Series) -> bool:
	return not col.any()

def is_unnamed(s: str) -> bool:
	if isinstance(s, str):
		return s.startswith('Unnamed: ')
	return False

def series_gen(df: pd.DataFrame, cols=None, allownan=False) -> pd.Series:
	if cols == None:
		cols = df.columns
	for col in cols:
		if not allownan and is_nan_series(df[col]) and is_unnamed(col):
			continue
		yield df[col]

def remove_and_order_columns(seriesDict : dict[str, pd.Series], col_order: list[str]) -> dict[str, pd.Series]:
	return {col : seriesDict[col] for col in col_order}

def clean_columns(df: pd.DataFrame) -> dict[str, pd.Series]:
	return {ser[0] : ser[1:] for ser in series_gen(df)}

def create_clean_table(df: pd.DataFrame, col_order: list[str]):
	return pd.DataFrame(remove_and_order_columns(clean_columns(df), col_order))