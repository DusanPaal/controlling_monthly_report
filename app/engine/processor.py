
"""
Description:
------------
The module provides support for data conversion,
preprocessing, aggregation and merging.

Version history:
----------------
1.0.20210515 - initial version
"""

import re
from io import StringIO
from logging import Logger, getLogger

import numpy as np
import pandas as pd
from pandas import DataFrame, Series

log: Logger = getLogger("master")


def _parse_amounts(vals: Series) -> Series:
	"""Converts string amounts in the SAP format to floating point literals."""

	parsed = vals.copy()
	parsed = parsed.str.replace(".", "", regex = False).str.replace(",", ".", regex = False)
	parsed = parsed.mask(parsed.str.endswith("-"), "-" + parsed.str.replace("-", ""))
	parsed = pd.to_numeric(parsed)

	return parsed

def _load_customer_data(branches_path: str, head_offices_path: str) -> DataFrame:
	"""Loads data detailing information of customers
	such as customer name, account, AR processor, etc...
	"""

	branches = pd.read_csv(
		branches_path,
		header = 0, sep = ";",
		encoding = "ansi",
		na_filter = False
	)

	head_offs = pd.read_csv(
		head_offices_path,
		header = 0, sep = ";",
		encoding = "iso8859_15",
		dtype = {
			"head_office": "UInt32",
			"country": "category",
			"Company_Code": "category",
			"type": "category"
		}
	)

	merged = pd.merge(branches, head_offs, how = "outer", on = "head_office")
	merged["branch_number"] = pd.to_numeric(merged["branch_number"]).astype("UInt32")
	merged["employee_id"] = pd.to_numeric(merged["employee_id"]).astype("UInt8")

	return merged

def _aggregate_data(preprocessed: DataFrame) -> DataFrame:
	"""Performs accounting data aggrgation based on defined parameters."""

	data = preprocessed.copy()

	# create a helper field with absolute DC amounts
	data["LC_Amount_ABS"] = data["LC_Amount"].abs()

	# categorize the absolute DC amounts into Deductions
	data["Deductions"] = pd.cut(
		x = data["LC_Amount_ABS"],
		right = True,
		labels = ["under 30", "30 - 50", "over 50"],
		bins = [0, 30, 50, 999999]
	).astype("category")

	# delete the helper field as this is no longer needed
	data.drop("LC_Amount_ABS", axis = 1, inplace = True)

	# copy fields to summarize in that new ones are created
	data["Deductions_Count"] = data["Deductions"]
	data["Deductions_Total"] = data["LC_Amount"]

	data["Customer_Number"].fillna(0, inplace = True)

	pivoted = data.pivot_table(
		index = [
			"Company_Code",
			"GL_Account",
			"Year",
			"Period",
			"Customer_Number",
			"Currency"
		],
		columns = "Deductions",
		values = ["Deductions_Total", "Deductions_Count"],
		aggfunc = {
			"Deductions_Total": lambda x: x.sum(),
			"Deductions_Count": lambda x: x.count()
		},
		fill_value = 0
	)

	# perform final data consolidation
	stacked = pivoted.stack().reset_index()

	# convert fields to appropriate data types
	stacked["Deductions_Count"] = stacked["Deductions_Count"].astype("uint16")
	stacked["Company_Code"] = stacked["Company_Code"].astype("category")

	return stacked

def assign_customers(
		compacted: DataFrame,
		branches_path: str,
		head_offices_path: str
	) -> DataFrame:
	"""Assigns a customer name to each item in the aggregated accounting data
	where a customer account is available. If no customer is identified for
	a given customer account, then the customer name value is left empty.

	Parameters:
	-----------
	compacted:
		aggregated accountig data.

	branches_path:
		DataFrame object containing customer data.

	Returns:
	--------
	A DataFrame object of aggregated accounting data with assigned customer names.
	"""

	if compacted.empty:
		raise ValueError("Input data contains no records!")

	aggregated = _aggregate_data(compacted)
	cust_data = _load_customer_data(branches_path, head_offices_path)

	updated_cust_num = pd.merge(
		aggregated,
		cust_data[[
			"branch_number",
			"Customer_Name",
			"Company_Code"
		]].dropna(axis = 0),
		how = "left",
		left_on = ["Customer_Number", "Company_Code"],
		right_on = ["branch_number", "Company_Code"]
	).drop(
		columns = "branch_number"
	)

	updated_hd_off = pd.merge(
		updated_cust_num,
		cust_data[[
			"head_office",
			"Customer_Name"
		]].drop_duplicates(),
		how = "left",
		left_on = "Customer_Number",
		right_on = "head_office"
	)

	mask = updated_hd_off["Customer_Name_x"].isna()
	updated_hd_off.loc[mask, "Customer_Name_x"] = updated_hd_off["Customer_Name_y"]
	dropped = updated_hd_off.drop("Customer_Name_y", axis = 1)
	renamed = dropped.rename({"Customer_Name_x": "Customer_Name"}, axis = 1)

	assert aggregated.shape[0] == renamed.shape[0], "Input and output data rows not equal!"

	return renamed

def convert_data(exp_data: str) -> list:
	"""Converts a list of plain FBL3N text data into panel datasets.

	Parameters:
	-----------
	exp_data:
		Plain text data exported form FBL3N.

	Returns:
	--------
	List of parsed panel data in the form of DataFrame object
	on success, None on failure. Each object in list represents
	data for a particular country.
	"""

	# define data header names
	header = [
		"Currency",
		"Company_Code",
		"GL_Account",
		"Year",
		"Period",
		"Document_Type",
		"Offsetting_Account",
		"Offsetting_Account_Type",
		"LC_Amount",
		"Text"
	]

	matches = re.findall(r"^\|\s*\w{3}\s*\|.*\|$", exp_data, re.M)

	raw_txt = "\n".join(matches)
	replaced = re.sub(r"^\|", "", raw_txt, flags = re.M)
	replaced = re.sub(r"\|$", "", replaced, flags = re.M)

	# clean data from reserved chars entered by users
	cleaned = re.sub("\"", "", replaced, flags = re.M)
	data = pd.read_csv(StringIO(cleaned), sep = "|", names = header, dtype = "string")

	# remove leading and trailing spaces from data vals
	for col in data:
		data[col] = data[col].str.strip()

	# convert fields to approprite data types
	data["LC_Amount"] = _parse_amounts(data["LC_Amount"].str.strip())
	data["GL_Account"] = data["GL_Account"].astype("uint64")
	data["Year"] = data["Year"].astype("uint16")
	data["Period"] = data["Period"].astype("uint8")
	data["Company_Code"] = data["Company_Code"].astype("category")
	data["Currency"] = data["Currency"].astype("category")
	data["Document_Type"] = data["Document_Type"].astype("category")
	data["Offsetting_Account"] = data["Offsetting_Account"].astype("uint64")
	data["Offsetting_Account_Type"] = data["Offsetting_Account_Type"].astype("category")

	return data

def compact_data(converted: list) -> DataFrame:
	"""Performs data preprocessing incl. concatenation, cleaning
	account extraction from text, and type conversion on
	the converted accounting data.

	Parameters:
	-------
	converted:
		A list of DataFrame objects that
		contain the converted accountig data.

	Returns:
	--------
	The peprocessed accounting data.
	"""

	if len(converted) == 0:
		raise ValueError("The input data contains no records!")

	# concatenate data for particular company codes
	data = pd.concat(converted, axis = 0).reset_index().drop("index", axis = 1)

	# clean concatenated data
	data["Text"].fillna("", inplace = True)

	# extract branch/head office number from text
	data = data.assign(
		Customer_Number = lambda x: x.Text.str.findall(r"\D+([1,4]\d{6})(?!\d)")
	)

	data["Customer_Number"] = data["Customer_Number"].apply(
		lambda x: np.uint64(x[0]) if len(x) >= 1 else pd.NA
	)

	# find records with missing debitor number (not found in the item text)
	# for all the offsetting accounts but the of cheque kind
	missing_cust_acc = data["Offsetting_Account"] != 48505240 & data["Customer_Number"].isna()
	data.loc[missing_cust_acc, "Customer_Number"] = data.loc[missing_cust_acc, "Offsetting_Account"]

	# cast to appropriate type where customer number exists
	exists_cust_acc = data["Customer_Number"].notna()
	data.loc[exists_cust_acc, "Customer_Number"] = data.loc[exists_cust_acc, "Customer_Number"].astype("uint32")

	return data
