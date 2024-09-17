# pylint: disable = C0103, W0706

"""
The controller.py represents the middle layer of the application design \n
and mediates communication between the top layer (app.py) and the \n
highly specialized modules situated on the bottom layer of the design \n
(fbl3n.py, processor.py, report.py sap.py).
"""

import shutil
import logging
import os
from datetime import datetime as dt
from datetime import timedelta
from glob import glob
from logging import Logger, config, getLogger
from os.path import basename, isfile, join
import yaml
from pandas import DataFrame
from win32com.client import CDispatch
from engine import fbl3n, processor, report, sap

FilePath = str
DirPath = str

log: Logger = getLogger("master")


# ====================================
# 		application configuration
# ====================================

def load_app_config(cfg_path: str) -> dict:
	"""Reads application configuration
	parameters from a file.

	Parameters:
	-----------
	cfg_path:
		Path to a yaml/yml file that contains
		application configuration parameters.

	Returns:
	--------
	Application configuration parameters.
	"""

	log.info("Loading application configuration ...")

	if not cfg_path.endswith((".yaml", ".yml")):
		raise ValueError("The configuration file not a YAML/YML type!")

	with open(cfg_path, encoding = "utf-8") as stream:
		content = stream.read()

	cfg = yaml.safe_load(content)
	log.info("Configuration loaded.")

	return cfg

def load_processing_rules(rules_path: str) -> dict:
	"""Reads data processing rules form a file.

	Parses application 'rules.yaml' file that contains
	country-specific data processing parameters.

	Parameters:
	-----------
	rules_path:
		Path to a yaml/yml file that contains
		the processing parameters (rules).

	Returns:
	--------
	Data processing rules.
	"""

	log.info("Loading processing rules ...")
	with open(rules_path, encoding = "utf-8") as stream:
		content = stream.read()

	rules = yaml.safe_load(content)

	for cocd in rules.copy():
		if not rules[cocd]["active"]:
			country = rules[cocd]["country"]
			log.warning(f"Processing of '{country}' disabled.")
			del rules[cocd]

	if len(rules) != 0:
		log.info("Rules loaded.")
	else:
		log.warning("No active country found.")

	return rules

# ====================================
# initialization of the logging system
# ====================================

def _compile_log_path(log_dir: str) -> str:
	"""Compiles the path to the log file
	by generating a log file name and then
	concatenating it to the specified log
	directory path."""

	date_tag = dt.now().strftime("%Y-%m-%d")
	nth = 0

	while True:
		nth += 1
		nth_file = str(nth).zfill(3)
		log_name = f"{date_tag}_{nth_file}.log"
		log_path = join(log_dir, log_name)

		if not isfile(log_path):
			break

	return log_path

def _read_log_config(cfg_path: str) -> dict:
	"""Reads logging configuration parameters from a yaml file."""

	# Load the logging configuration from an external file
	# and configure the logging using the loaded parameters.

	if not isfile(cfg_path):
		raise FileNotFoundError(f"The logging configuration file not found: '{cfg_path}'")

	with open(cfg_path, encoding = "utf-8") as stream:
		content = stream.read()

	return yaml.safe_load(content)

def _update_log_filehandler(log_path: str, logger: Logger) -> None:
	"""Changes the log path of a logger file handler."""

	prev_file_handler = logger.handlers.pop(1)
	new_file_handler = logging.FileHandler(log_path)
	new_file_handler.setFormatter(prev_file_handler.formatter)
	logger.addHandler(new_file_handler)

def _print_log_header(logger: Logger, header: list, terminate: str = "\n") -> None:
	"""Prints header to a log file."""

	for nth, line in enumerate(header, start = 1):
		if nth == len(header):
			line = f"{line}{terminate}"
		logger.info(line)

def _remove_old_logs(logger: Logger, log_dir: str, n_days: int) -> None:
	"""Removes old logs older than the specified number of days."""

	old_logs = glob(join(log_dir, "*.log"))
	n_days = max(1, n_days)
	curr_date = dt.now().date()

	for log_file in old_logs:
		log_name = basename(log_file)
		date_token = log_name.split("_")[0]
		log_date = dt.strptime(date_token, "%Y-%m-%d").date()
		thresh_date = curr_date - timedelta(days = n_days)

		if log_date < thresh_date:
			try:
				logger.info(f"Removing obsolete log file: '{log_file}' ...")
				os.remove(log_file)
			except PermissionError as exc:
				logger.error(str(exc))

def configure_logger(log_dir: str, cfg_path: str, *header: str) -> None:
	"""Configures application logging system.

	Parameters:
	-----------
	log_dir:
		Path t the directory to store the log file.

	cfg_path:
		Path to a yaml/yml file that contains
		application configuration parameters.

	header:
		A sequence of lines to print into the log header.
	"""

	log_path = _compile_log_path(log_dir)
	log_cfg = _read_log_config(cfg_path)
	config.dictConfig(log_cfg)
	logger = logging.getLogger("master")
	_update_log_filehandler(log_path, logger)
	if header is not None:
		_print_log_header(logger, list(header))
	_remove_old_logs(logger, log_dir, log_cfg.get("retain_logs_days", 1))


# ====================================
# 		Management of SAP connection
# ====================================

def connect_to_sap(system: str) -> CDispatch:
	"""Creates connection to the SAP GUI scripting engine.

	Parameters:
	-----------
	system:
		The SAP system to use for connecting to the scripting engine.

	Returns:
	--------
	A SAP `GuiSession` object that represents active user session.
	"""

	log.info("Connecting to SAP ...")
	sess = sap.connect(system)
	log.info("Connection created.")

	return sess

def disconnect_from_sap(sess: CDispatch) -> None:
	"""Closes connection to the SAP GUI scripting engine.

	Parameters:
	-----------
	sess:
		An SAP `GuiSession` object (wrapped in the `win32:CDispatch` class)
		that represents an active user SAP GUI session.
	"""

	log.info("Disconnecting from SAP ...")
	sap.disconnect(sess)
	log.info("Connection to SAP closed.")

def _calculate_date_range() -> tuple:
	"""Calculates first and last dates of data
	range for which FBL3N data will be exported.
	"""

	date = dt.date(dt.now())
	last_day_prev_mon = date.replace(day = 1) - timedelta(1)
	first_day_prev_mon = last_day_prev_mon.replace(day = 1)

	return (first_day_prev_mon, last_day_prev_mon)


# ====================================
# 		data export and processing
# ====================================

def export_fbl3n_data(
		sess: CDispatch, layout: str,
		temp_dir: str, rules: dict,
		max_attempts: int = 3
	) -> DataFrame:
	"""Exports data from GL accounts into a local file.

	Parameters:
	-----------
	sess:
		A SAP `GuiSession` object.

	layout:
		Name of the FBL3N layout used
		for formatting of the accounting data.

	temp_dir:
		Path to the directory to store temporary application files.

	rules:
		Country-specific data processing rules.

	max_attempts:
		Number of attempts to try exporting the data before giving up.

	Returns:
	--------
	A list of `pandas.DataFrame` objects that contain the exported data.
	"""

	log.info("Starting FBL3N ...")
	fbl3n.start(sess)
	log.info("The FBL3N has been started.")

	from_date, to_date = _calculate_date_range()
	result = []

	for cocd in rules:

		country = rules[cocd]["country"]
		accs = rules[cocd]["accounts"]

		exp_path = join(temp_dir, "fbl3n_exp.txt")
		log.info(f"Exporting data for country: '{country}'")
		nth = 0

		while nth < max_attempts:
			try:
				data = fbl3n.export(exp_path, accs, cocd, from_date, to_date, layout)
			except fbl3n.SapConnectionLostError as exc:
				log.error(exc)
				log.info(f"Attempt # {nth} of {max_attempts} to handle the error ...")
				nth += 1
			except fbl3n.DataExportError:
				raise
			else:
				nth = 0
				break

		if nth != 0:
			raise RuntimeError("Attempts to handle the SapConnectionLostError failed!")

		log.info("Converting exported data ...")
		converted = processor.convert_data(data)
		result.append(converted)
		log.info("Data converted successfully.")

	log.info("Closing FBL3N ...")
	fbl3n.close()
	log.info("The FBL3N has been closed.")

	return result

def process_fbl3n_data(exported: list, branches_path: str, head_offices_path: str) -> dict:
	"""Manages processing of FBL3N data exported into a plain text file.

	Parameters:
	-----------
	exported:
		A list of pandas.DataFrame objects that contain the exported FBL3N data.

	branches_path:
		Path to the file where information about branches is stored.

	head_offices_path:
		Path to the file where information about head office accounts is stored.

	Returns:
	--------
	The result of data processing:
		"preprocessed": `pandas.DataFrame`
			Preprocessed data.
		"updated": `pandas.DataFrame`
			Preprocessed data with assigned customer names stored in the "Customer_Name" field.
	"""

	log.info("Preprocessig data ...")
	compacted = processor.compact_data(exported)
	log.info("Data preprocessed successfully.")

	log.info("Updating data on customer names ...")
	updated = processor.assign_customers(compacted, branches_path, head_offices_path)
	log.info("Data updated successfully.")

	return {"compacted": compacted,	"updated": updated}


# ====================================
# 	Reporting of processing output
# ====================================

def report_output(
		temp_dir: str,
		pivot_path: str,
		compacted: DataFrame,
		updated: DataFrame,
		report_cfg: dict
	) -> None:
	"""Manages creating and uploding of the user report.

	Parameters:
	-----------
	temp_dir:
		Path to the directory to store temporary application files.

	pivot_path:
		Path to the xlsm file with a VBA script
		that generates the pivot table of the report.

	compacted:
		Compacted FBL5N data.

	updated:
		Raw FBL5N data containing customer info.

	report_cfg:
		Application 'reporting' configuration parameters.
	"""

	calendar_month = dt.now().strftime("%m")
	calendar_year = dt.now().strftime("%Y")

	report_name = report_cfg["name"].replace("$calendar_year$", calendar_year)
	report_name = report_name.replace("$calendar_month$", calendar_month)

	if not report_name.lower().endswith(".xlsx"):
		report_name = "".join([report_name, ".xlsx"])

	src_report_path = join(temp_dir, report_name)

	src_report_path = report.generate_excel_report(
		updated, compacted, src_report_path, pivot_path, temp_dir,
		exported_sht_name = report_cfg["exported_datasheet_name"],
		processed_sht_name = report_cfg["processed_datasheet_name"],
		pivotted_sht_name = report_cfg["pivotted_datasheet_name"]
	)

	upload_dir = report_cfg["upload_dir"]
	dst_report_path = join(upload_dir, basename(src_report_path))
	log.info(f"Uploading report: '{src_report_path}' -> '{dst_report_path}' ...")

	try:
		shutil.move(src_report_path, upload_dir)
	except shutil.Error as exc:
		log.error(exc)
		log.warning("The existing file will be overwrittten.")
		os.remove(dst_report_path)
		shutil.move(src_report_path, upload_dir)
