# pylint: disable = E0110

"""The module creates user reports from the processing output."""

import logging
import os
from logging import Logger
from os.path import basename, exists, join
from zipfile import ZipFile

import pandas as pd
from pandas import DataFrame, ExcelWriter

HEADER_ROW_IDX = 1

log: Logger = logging.getLogger("master")

def _extract_macro(xlsm_path: str, bin_filepath: str) -> None:
	"""Extracts binary content with VBA macro from an xlsm report template file."""

	if not exists(xlsm_path):
		raise FileNotFoundError(f"Path to the XLSM file not found: '{xlsm_path}'")

	bin_relpath = "".join(["xl/", basename(bin_filepath)])

	with ZipFile(xlsm_path) as stream:
		vba_bin_data =  stream.read(bin_relpath)

	with open(bin_filepath, "wb") as stream:
		stream.write(vba_bin_data)

def _get_col_width(data: DataFrame, data_fld_name: str) -> int:
	"""Returns excel column width calculated as
	the maximum count of characters contained
	in the column name and column data strings.
	"""

	data_fld_vals = data[data_fld_name]
	vals = data_fld_vals.astype("string").dropna().str.len()
	vals = list(vals)
	vals.append(len(str(data_fld_name)))

	return max(vals)

def generate_excel_report(
		processed: DataFrame,
		exported: DataFrame,
		report_path: str,
		vba_path: str,
		temp_path: str,
		**sheet_names
	) -> None:
	"""Creates an Excel user report.

	Parameters:
	-----------
	processed:
		The data processing result.
	
	exported:
		The raw data export.

	report_path:
		Path to the report Excel file.

	vba_path:
		Path to the xlsm file containing VBA
		macro that creates the summary pivot table.

	temp_path:
		Path to application temporary data folder.

	**sheet_names:
		Names of report sheets to which data will be written.
	"""

	if processed.empty:
		raise ValueError("The processed data contains no records!")

	if exported.empty:
		raise ValueError("The exported data contains no records!")

	# post-process data
	processed.loc[processed["Customer_Number"] == 0, "Customer_Number"] = pd.NA
	exported.loc[exported["Customer_Number"] == 0, "Customer_Number"] = pd.NA

	# customize data layout by reording fields
	processed_header = [
		"Company_Code",
		"GL_Account",
		"Period",
		"Year",
		"Customer_Number",
		"Customer_Name",
		"Currency",
		"Deductions",
		"Deductions_Count",
		"Deductions_Total"
	]

	raw_header = [
		"Company_Code",
		"Year",
		"Period",
		"Document_Type",
		"GL_Account",
		"Customer_Number",
		"Currency",
		"LC_Amount",
		"Text",
		"Offsetting_Account",
		"Offsetting_Account_Type"
	]

	processed = processed[processed_header]
	exported = exported[raw_header]

	# print datasets to dedicated sheets of a workbook
	exp_sht_name = sheet_names["exported_sht_name"]
	proc_sht_name = sheet_names["processed_sht_name"]
	pivot_sht_name = sheet_names["pivotted_sht_name"]

	with ExcelWriter(report_path, engine = "xlsxwriter") as wrtr:

		# format headers
		processed.columns = processed.columns.str.replace("_", " ", regex=False)
		exported.columns = exported.columns.str.replace("_", " ", regex=False)

		exported.to_excel(wrtr, index = False, sheet_name = exp_sht_name)
		processed.to_excel(wrtr, index = False, sheet_name = proc_sht_name)

		processed.columns = processed.columns.str.replace(" ", "_", regex=False)
		exported.columns = exported.columns.str.replace(" ", "_", regex=False)

		# format data
		report = wrtr.book  # pylint: disable=E1101
		proc_data_sht = wrtr.sheets[proc_sht_name]
		exp_data_sht = wrtr.sheets[exp_sht_name]
		money_fmt = report.add_format({"num_format": "#,##0.00", "align": "center"})
		general_fmt = report.add_format({"align": "center"})

		# apply formats to 'Summary' data sheet
		for idx, col_name in enumerate(processed):
			col_fmt = money_fmt if col_name == "Deductions_Total" else general_fmt
			col_width = _get_col_width(processed, col_name) + 2
			proc_data_sht.set_column(idx, idx, col_width, col_fmt)

		# apply formats to 'Source' data sheet
		for idx, col_name in enumerate(exported):
			col_fmt = money_fmt if col_name == "LC_Amount" else general_fmt
			col_width = _get_col_width(exported, col_name) + 2
			exp_data_sht.set_column(idx, idx, col_width, col_fmt)

		# extract vba file that contains macro for
		# pivot table creation and add it to the report
		bin_filepath = join(temp_path, "vbaProject.bin")
		_extract_macro(vba_path, bin_filepath)
		report.add_vba_project(bin_filepath)
		pivotted = report.add_worksheet(pivot_sht_name)

		# adjust cell width and height to fit the button
		pivotted.set_row(0, 22)
		pivotted.set_column("A:A", 12)

		# Add 'Create' button to the sheet
		# containing  the pivot table
		pivotted.insert_button("A1", {
			"macro": "CreatePivotTable",
			"caption": "Create",
			"width": 90,
			"height": 30
		})

		# freeze headers on data sheets
		proc_data_sht.freeze_panes(HEADER_ROW_IDX, 0)
		exp_data_sht.freeze_panes(HEADER_ROW_IDX, 0)

	# rename to xlsm in order for EXCEL to open the report
	new_report_path = report_path.replace("xlsx", "xlsm")

	try:
		os.rename(report_path, new_report_path)
	except FileExistsError as exc:
		log.error(exc)
		log.warning("The existing file will be overwrittten.")
		os.remove(new_report_path)
		os.rename(report_path, new_report_path)

	try:
		os.remove(bin_filepath)
	except (PermissionError, FileNotFoundError) as exc:
		log.error(exc)

	return new_report_path
