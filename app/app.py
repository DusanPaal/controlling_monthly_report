# pylint: disable = C0103

"""
Description:
------------
The 'app.py' module represents the main script of the application that contains
program entry procedure called at the script import.

The 'FBL3N monthly report' application automates generating of a monthly controlling
report which 
An excel report file is generated and uploaded to a folder specified by the users.

Version history:
----------------
1.0.20210515 - initial version
1.0.20220610 - refactored version
"""

from os.path import join
from datetime import datetime as dt
import logging
import sys
from engine import controller

log = logging.getLogger("master")

def main() -> int:
	"""
	Program entry point.

	Returns:
	--------
	An integer representing the program's completion state:
	- 0: Program successfully completes.
	- 1: Program fails in while configuring the loging system.
	- 2: Program fails in the initialization phase.
	- 3: Program fails in the processing phase.
	"""

	app_dir = sys.path[0]
	log_dir = join(app_dir, "logs")
	temp_dir = join(app_dir, "temp")
	app_cfg_path = join(app_dir, "app_config.yaml")
	log_cfg_path = join(app_dir, "log_config.yaml")
	rules_path = join(app_dir, "rules.yaml")
	pivot_path = join(app_dir, "data", "pivotting", "macro.xlsm")
	branches_path = join(app_dir, "data", "customers", "branches.csv")
	head_office_path = join(app_dir, "data", "customers", "head_offices.csv")
	curr_date = dt.now().strftime("%d-%b-%Y")

	try:
		controller.configure_logger(
			log_dir, log_cfg_path,
			"Application name: Controllig FBL3N monthly report",
			"Application version: 1.0.20220610",
			f"Log date: {curr_date}")
	except Exception as exc:
		print(exc)
		print("CRITICAL: Unhandled exception while trying to configuring the logging system!")
		return 1

	try:
		log.info("=== Initialization START ===")
		cfg = controller.load_app_config(app_cfg_path)
		rules = controller.load_processing_rules(rules_path)

		if len(rules) == 0:
			log.info("=== Initialization END ===\n")
			return 0

		sess = controller.connect_to_sap(cfg["sap"]["system"])
		log.info("=== Initialization END ===\n")
	except Exception as exc:
		log.critical(exc)
		return 2

	try:

		log.info("=== Data export START ===")
		exported = controller.export_fbl3n_data(
			sess, cfg["data"]["fbl3n_layout"], temp_dir, rules)
		log.info("=== Data export END ===\n")

		log.info("=== Processing START ===")
		result = controller.process_fbl3n_data(
			exported, branches_path, head_office_path)
		log.info("=== Processing END ===\n")

		log.info("=== Reporting START ===")
		controller.report_output(
			temp_dir, pivot_path, result["compacted"],
			result["updated"], cfg["reports"])
		log.info("=== Reporting END ===\n")

	except Exception as exc:
		log.exception(exc)
		log.critical("Processing terminated due to an uhandled exception!")
		return 3
	finally:
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(sess)
		log.info("=== Cleanup END ===\n")

	return 0

if __name__ == "__main__":
	exit_code = main()
	log.info(f"=== System shutdown with return code: {exit_code} ===")
	sys.exit(exit_code)
