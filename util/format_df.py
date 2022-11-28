import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import Callable

@dataclass(frozen=True)
class WriterInfo:
	writer: pd.ExcelWriter
	cleaned_df: pd.DataFrame
	sheet_name: str
	startrow: int
	startcol: int
	col_order: list[str]

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



def by_longest_word_in_colname(colname: str):
	return len(
				max(
					colname.split(' '), 
					key=len)
				)
