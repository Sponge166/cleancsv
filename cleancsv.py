import pandas as pd
import numpy as np
from typing import Callable
from dataclasses import dataclass
import argparse
from pathlib import Path

@dataclass(frozen=True)
class WriterInfo:
	writer: pd.ExcelWriter
	cleaned_df: pd.DataFrame
	sheet_name: str
	startrow: int
	startcol: int
	col_order: list[str]

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

def apply_format_to_entire_col(writerinfo: WriterInfo, 
							   key: Callable[str, bool], 
							   col_format):
	worksheet = writerinfo.writer.sheets[writerinfo.sheet_name]

	for i, col in enumerate(writerinfo.col_order):
		if not key(col):
			continue

		startrow_up1 = writerinfo.startrow + 1
		# adding the extra 1 to account for the index column
		column_idx = i + startrow_up1 + 1
		for row_idx in range(startrow_up1, len(writerinfo.cleaned_df)+startrow_up1):
			val = writerinfo.cleaned_df.iloc[row_idx-(startrow_up1), i]
			if isinstance(val, float) and np.isnan(val):
				val = ''
			worksheet.write_string(row_idx, column_idx, val, col_format)


def highlight_cols(writerinfo: WriterInfo, 
				   key: Callable[str, bool],
				   color: str='#fff2cc'):
	highlight = writerinfo.writer.book.add_format({'bg_color': color})

	apply_format_to_entire_col(writerinfo, key, highlight)

def widen_cols(writerinfo: WriterInfo, 
			   colnames: set[str], 
			   width_func: Callable[str, int],
			   wordwrapcolname=False,
			   width_delta: int=3):
	worksheet = writerinfo.writer.sheets[writerinfo.sheet_name]

	for i, col in enumerate(writerinfo.col_order):
		if col not in colnames:
			continue
			
		column_idx = i + writerinfo.startcol + 1
		# adding the extra 1 to account for the index column
		if wordwrapcolname:
			text_wrap = writerinfo.writer.book.add_format({'text_wrap':1, 'bg_color':'#cccccc', 'bold': True})
			worksheet.write_string(writerinfo.startrow, column_idx, col, text_wrap)

		# column_title_font_size = worksheet.table[writerinfo.startrow][column_idx].format.__dict__['font_size']

		# print(column_title_font_size)

		width = width_func(col)
		width += width_delta

		worksheet.set_column(column_idx, column_idx, width)

def freeze(writerinfo: WriterInfo):
	writerinfo.writer.sheets[writerinfo.sheet_name].freeze_panes(writerinfo.startrow+1, writerinfo.startcol+2)

def verify_dest(dest: str | Path) -> Path | None:
	p = Path(dest)
	if p.suffix not in {'.xlsx', '.xls', '.xlsm'}:
		raise ValueError('Destination file must be an excel file: ends with ".xlsx" or ".xls" or ".xlsm"')
	return p

def verify_source(source: str | Path) -> Path | None:
	p = Path(source)
	if p.suffix != '.csv':
		raise ValueError('Source file must be a csv file: ends with ".csv"')
	return p

def main():

	COL_ORDER = ['ZoomInfo Company ID', 'Company Name', 'Revenue (in 000s USD)', 'Revenue Range (in USD)', 'Employees', 
			'Number of Locations', 'Company City', 'Company Zip Code', 'Website', 'Founded Year', 
			'Company HQ Phone', 'ZoomInfo Company Profile URL', 'LinkedIn Company Profile URL', 
			'Facebook Company Profile URL', 'Twitter Company Profile URL', 'Primary Industry', 
			'Primary Sub-Industry', 'All Industries', 'All Sub-Industries', 'Industry Hierarchical Category', 
			'Secondary Industry Hierarchical Category', 'Ownership Type', 'Business Model', 
			'Certified Active Company', 'Certification Date', 'Total Funding Amount (in 000s USD)', 
			'Recent Funding Amount (in 000s USD)', 'Recent Funding Round', 'Recent Funding Date', 
			'Recent Investors', 'All Investors', 'Full Address', 'Company Is Acquired', 
			'Company ID (Ultimate Parent)', 'Entity Name (Ultimate Parent)', 'Relationship (Immediate Parent)']

	class my_ArgumentParser(argparse.ArgumentParser):
		def exit(self, status=0, message=None):
			print('\n\tArguments are denoted by dashes "-".\n\tIf either your source file or destination file contain dashes surround them in double quotes\n\tsyntax: cleancsv "source" "dest" [-args]\n')
			super().exit(status, message)

	parser = my_ArgumentParser(description="Clean csv from source and save in destination")
	parser.add_argument('source', type=str)
	parser.add_argument('dest', type=str)
	parser.add_argument('-sn', default='newly_cleaned', type=str, required=False, help="-sn [desired sheet_name]")
	parser.add_argument('-sr', default=1, type=int, required=False, help="-sr [desired start_row]")
	parser.add_argument('-sc', default=2, type=int, required=False, help="-sc [desired start_col]")

	args = parser.parse_args()
	source = verify_source(args.source)
	dest = verify_dest(args.dest)

	unclean_df = pd.read_csv(source)
	cleaned_df = create_clean_table(unclean_df, COL_ORDER)

	writer = pd.ExcelWriter(dest, engine='xlsxwriter')
	writerinfo = WriterInfo(
		writer,
		cleaned_df,
		args.sn,
		args.sr,
		args.sc,
		COL_ORDER
		)
	
	cleaned_df.to_excel(writer, 
		startrow=writerinfo.startrow, 
		startcol=writerinfo.startcol, 
		sheet_name=writerinfo.sheet_name
		)

	highlight_cols(
		writerinfo, 
		lambda x: x.lower().endswith('url')
		)

	# setting width to the width of the column name
	by_col = lambda colname : len(colname)
	widen_cols(
		writerinfo, 
		set(COL_ORDER), 
		by_col, 
		True
		)

	# setting width to the width of the longest word in the column title
	by_long_col_word = lambda colname : len(max(colname.split(' '), key=len))
	cols = {'Total Funding Amount (in 000s USD)',
			'Recent Funding Amount (in 000s USD)',
			'Certified Active Company',
			'Business Model',
			'Industry Hierarchical Category',
			'Secondary Industry Hierarchical Category',
			'Company Name',
			'Founded Year',
			'Number of Locations',
			'Company Zip Code'}

	widen_cols(
		writerinfo, 
		cols, 
		by_long_col_word,
		width_delta=0
		)

	# setting width to the width of the longest value in the column
	by_longest_val = lambda colname : len(max([str(val) for val in writerinfo.cleaned_df[colname]], key=len)) * (10/11)
	# * (10/11) bc the items in each column have a font size of 10 and column width value is relative to size 11 of the default font
	cols = {'ZoomInfo Company ID',
			'Company Name',
			'Revenue (in 000s USD)',
			'Revenue Range (in USD)',
			'Employees',
			'Company City'}

	widen_cols(
		writerinfo,
		cols,
		by_longest_val
		)

	widen_cols(writerinfo, {'Website'}, lambda _ : len('ZoomInfo Company Profile URL'))

	freeze(writerinfo)

	writer.close()

	fairwell_message = '''
	Your csv has been successfully cleaned and converted to an excel file!
	Did this save you a lot of time? If so please consider tipping the developer.
	Venmo: @ClarkMattoon'''

	print(fairwell_message)
main()
