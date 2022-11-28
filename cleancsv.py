import pandas as pd
import argparse
from pathlib import Path
from util.clean_df import create_clean_table
from util.format_df import *

def verify_dest(dest: str | Path, source: str | Path) -> Path | None:
	if dest is None:
		return Path(source.parent, f'{source.stem}_cleaned.xlsx')

	p = Path(dest)
	if not p.suffix:
		p = Path(p, f'{source.stem}_cleaned.xlsx')
	elif p.suffix not in {'.xlsx', '.xls', '.xlsm'}:
		raise ValueError('Destination file must be an excel file: ends with ".xlsx" or ".xls" or ".xlsm"')
	return p

def verify_source(source: str | Path) -> Path | None:
	p = Path(source)
	if p.suffix != '.csv':
		raise ValueError('Source file must be a csv file: ends with ".csv"')
	return p

class my_ArgumentParser(argparse.ArgumentParser):
		def exit(self, status=0, message=None):
			print('\n\tArguments are denoted by dashes "-".\n\tIf either your source file or destination folder/file contain dashes surround them in double quotes\n\tsyntax: cleancsv "source" [-d "dest"] [-args]\n')
			super().exit(status, message)

def createParser():
	parser = my_ArgumentParser(description="Clean csv from source and save in destination")
	parser.add_argument('source', type=str)
	parser.add_argument('-d', '--dest', default=None, type=str, required=False, help="-d, --dest [desired destination]")
	parser.add_argument('-sn', default='newly_cleaned', type=str, required=False, help="-sn [desired sheet_name]")
	parser.add_argument('-sr', default=1, type=int, required=False, help="-sr [desired start_row]")
	parser.add_argument('-sc', default=2, type=int, required=False, help="-sc [desired start_col]")

	return parser

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

	parser = createParser()

	args = parser.parse_args()
	source = verify_source(args.source)
	dest = verify_dest(args.dest, source)

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
	widen_cols(
		writerinfo, 
		set(COL_ORDER), 
		len,
		True
		)

	# setting width to the width of the longest word in the column title
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
		by_longest_word_in_colname,
		width_delta=0
		)

	# setting width to the width of the longest value in the column
	by_longest_val_in_col = lambda colname : \
		len(max((str(val) for val in writerinfo.cleaned_df[colname]), key=len)) * (10/11)
	# * (10/11) bc the items in each column have a font size of 10 
	# and column width value is relative to size 11 of the default font

	cols = {'ZoomInfo Company ID',
			'Company Name',
			'Revenue (in 000s USD)',
			'Revenue Range (in USD)',
			'Employees',
			'Company City'}

	widen_cols(
		writerinfo,
		cols,
		by_longest_val_in_col
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
