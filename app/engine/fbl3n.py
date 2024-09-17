# pylint: disable = C0103, C0123, W0603, W0703, W1203

"""
Description:
------------
The module automates data export from G/L accounts to a file.

Version history:
----------------
1.0.20210720 - initial version
1.0.20220218 - removed dymamic layout creation upon data load.
			   Data layouts will now be applied by entering a layout name
			   in the transaction main search mask.
"""

import os
from datetime import date
from logging import Logger, getLogger
from os.path import exists, split
from pyperclip import copy as copy_to_clipboard
from win32com.client import CDispatch

FilePath = str

_sess = None
_main_wnd = None
_stat_bar = None

log: Logger = getLogger("master")

_virtual_keys = {
	"Enter":        0,
	"F3":           3,
	"F8":           8,
	"F9":           9,
	"CtrlS":        11,
	"F12":          12,
	"ShiftF4":      16,
	"ShiftF12":     24,
	"CtrlF1":       25,
	"CtrlF8":       32,
	"CtrlShiftF6":  42
}

# custom exceptions and warnings
class NoItemsFoundWarning(Warning):
	"""Raised when no items are found on account(s)
	using the specified selection criteria.
	"""

class UninitializedModuleError(Exception):
	"""Raised when attempting to use a procedure
	before starting the transaction.
	"""

class SapConnectionLostError(Exception):
	"""Raised when the connection to SAP is lost."""

class FolderNotFoundError(Exception):
	"""Raised when a folder is reruired but doesn't exist."""

class DataExportError(Exception):
	"""Raised when data export fails."""

def _clear_clipboard() -> None:
	"""Clears the contents of the clipboard."""
	copy_to_clipboard("")

def _press_key(name: str) -> None:
	"""Simulates pressing a keyboard button."""
	_main_wnd.SendVKey(_virtual_keys[name])

def _is_popup_dialog() -> bool:
	"""Checks if the active window is a popup dialog window."""
	return _sess.ActiveWindow.type == "GuiModalWindow"

def _close_popup_dialog(confirm: bool) -> None:
	"""Confirms or declines a pop-up dialog."""

	if _sess.ActiveWindow.text == "Information":
		if confirm:
			_press_key("Enter") # confirm
		else:
			_press_key("F12")   # decline
		return

	btn_caption = "Yes" if confirm else "No"

	for child in _sess.ActiveWindow.Children:
		for grandchild in child.Children:
			if grandchild.Type != "GuiButton":
				continue
			if btn_caption != grandchild.text.strip():
				continue
			grandchild.Press()
			return

def _set_company_code(val: str) -> None:
	"""Enters company code value into the 'Company code'
	field located on the transaction main window.
	"""

	if not (len(val) == 4 and val.isnumeric()):
		raise ValueError(f"Invalid company code: '{val}'!")

	if _main_wnd.findAllByName("SD_BUKRS-LOW", "GuiCTextField").count > 0:
		_main_wnd.findByName("SD_BUKRS-LOW", "GuiCTextField").text = val
	elif _main_wnd.findAllByName("SO_WLBUK-LOW", "GuiCTextField").count > 0:
		_main_wnd.findByName("SO_WLBUK-LOW", "GuiCTextField").text = val

def _set_layout(val: str) -> None:
	"""Enters layout name into the 'Layout' field
	located on the transaction main search window.
	"""
	_main_wnd.findByName("PA_VARI", "GuiCTextField").text = val

def _set_accounts(vals: list) -> None:
	"""Opens 'G/L account' listbox and enters
	the account numbers into the search subwindow.
	"""

	# remapping to str is needed since
	# accounts may be passed as integers
	accs = list(map(str, vals))

	# open selection table for company codes
	_main_wnd.findByName("%_SD_SAKNR_%_APP_%-VALU_PUSH", "GuiButton").press()

	_press_key("ShiftF4")   				# clear any previous values
	copy_to_clipboard("\r\n".join(accs))    # copy accounts to clipboard
	_press_key("ShiftF12")  				# confirm selection
	_clear_clipboard()
	_press_key("F8")        				# confirm

def _set_posting_dates(first: date, last: date) -> None:
	"""Enters start and end posting dates in the transaction main window
	that define the date range for which accounting data will be loaded.
	"""

	sap_date_format = "%d.%m.%Y"
	date_from = first.strftime(sap_date_format)
	date_to = last.strftime(sap_date_format)

	log.debug(f"Export date range: {date_from} - {date_to}")

	_main_wnd.FindByName("SO_BUDAT-LOW", "GuiCTextField").text = date_from
	_main_wnd.FindByName("SO_BUDAT-HIGH", "GuiCTextField").text = date_to

def _toggle_worklist(activate: bool) -> None:
	"""Activates or deactivates the 'Use worklist'
	option in the transaction main search window.
	"""

	used = _main_wnd.FindAllByName("PA_WLSAK", "GuiCTextField").Count > 0

	if (activate or used) and not (activate and used):
		_press_key("CtrlF1")

def _set_line_items_selection(status: str) -> None:
	"""Sets line item selection mode by item status."""

	obj_name = None

	if status == "open":
		obj_name = "X_OPSEL"
	elif status == "cleared":
		obj_name = "X_CLSEL"
	elif status =="all":
		obj_name = "X_AISEL"

	assert obj_name is not None, (f"Unrecognized selection status: '{status}'! "
	"Choose one option from: 'open', 'closed', 'all'.")

	_main_wnd.FindByName(obj_name, "GuiRadioButton").Select()

def _export_to_file(file: FilePath) -> None:
	"""Exports data to a local file."""

	folder_path, file_name = split(file)

	if not exists(folder_path):
		raise FolderNotFoundError(
			"The export folder not found at the "
			f"path specified: '{folder_path}'!")

	# open local data file export dialog
	_press_key("F9")

	# set plain text data export format and confirm
	_sess.FindById("wnd[1]").FindAllByName("SPOPLI-SELFLAG", "GuiRadioButton")(0).Select()
	_press_key("Enter")

	# enter data export file name, folder path and encoding
	# then click 'Replace' an existing file button
	folder_path = "".join([folder_path, "\\"])
	utf_8 = "4120"

	_sess.FindById("wnd[1]").FindByName("DY_PATH", "GuiCTextField").text = folder_path
	_sess.FindById("wnd[1]").FindByName("DY_FILENAME", "GuiCTextField").text = file_name
	_sess.FindById("wnd[1]").FindByName("DY_FILE_ENCODING", "GuiCTextField").text = utf_8
	_press_key("CtrlS")

def _read_exported_data(file_path: str) -> str:
	"""Reads exported FBL3N data from the text file."""

	with open(file_path, encoding = "utf-8") as stream:
		text = stream.read()

	return text

def _check_prerequisities():
	"""Verifies that the prerequisites
	for using the module are met."""

	if _sess is None:
		raise UninitializedModuleError(
			"Uninitialized module! Use the start() "
			"procedure to run the transaction first!")

def start(sess: CDispatch) -> None:
	"""Starts the FBL5N transaction.

	If the FBL5N has already been started,
	then the running transaction will be restarted.

	Parameters:
	-----------
	sess:
		An SAP `GuiSession` object (wrapped in
		the win32:CDispatch class)that represents
		an active user SAP GUI session.
	"""

	global _sess
	global _main_wnd
	global _stat_bar

	if sess is None:
		raise UnboundLocalError("Argument 'sess' is unbound!")

	# close the transaction
	# if it is already running
	close()

	_sess = sess
	_main_wnd = _sess.findById("wnd[0]")
	_stat_bar = _main_wnd.findById("sbar")
	_sess.startTransaction("FBL3N")

def close() -> None:
	"""Closes a running FBL5N transaction.

	Attempt to close the transaction that has not been \n
	started by the `start()` procedure is ignored.
	"""

	global _sess
	global _main_wnd
	global _stat_bar

	if _sess is None:
		return

	_sess.EndTransaction()

	if _is_popup_dialog():
		_close_popup_dialog(confirm = True)

	_sess = None        # type: ignore
	_main_wnd = None    # type: ignore
	_stat_bar = None    # type: ignore

def export(
		file: FilePath, accs: list,
		cocd: str, from_day: date,
		to_day: date, layout: str = ""
	) -> str:
	"""Exports item data from GL accounts.

	A `NoItemsFoundWarning` warning is raised
	if no items are found for the given selection criteria.

	An `DataExportError` exception is raised
	the attempt to expot the accounting data fails.

	An `SapConnectionLostError` exception is raised
	if the connection to SAP is lost due to an error.

	Prerequisities:
	---------------
	The FBL3N must be started by calling the `start()` procedure.

	Attempt to use the procedure when FBL3N has not been started \n
	results in the `UninitializedModuleError` exception.

	Parameters:
	-----------
	file:
		Path to a temporary .txt file to which the data is exported \n
		before reading. If the file path refers to an invalid folder, \n
		then a `FolderNotFoundError` exception is raised. The file is 
		removed once the data reading is complete.

	accs:
		GL accounts from which data is exported.

	cocd:
		Company code to which the accounts are assigned. \n
		A valid company code is a 4-digit string (e.g. '0075').

	from_day:
		Posting date from which accounting data is loaded.

	to_day:
		Posting date up to which accounting data is loaded.

	layout:
		Name of the layout that defines the format of the loaded data. \n
		By default, no specific layout name is used.

	Returns:
	--------
	The exported data as plain text.
	"""

	_check_prerequisities()

	_toggle_worklist(activate = False)
	_set_company_code(cocd)
	_set_layout(layout)
	_set_accounts(accs)
	_set_line_items_selection(status = "all")
	_set_posting_dates(from_day, to_day)
	_press_key("F8") # load item list

	try: # SAP crash can be caught only after next statement following item loading
		msg = _stat_bar.Text
	except Exception as exc:
		raise SapConnectionLostError("Connection to SAP lost!") from exc

	if "No items selected" in msg:
		raise NoItemsFoundWarning("No items found for the given selection criteria!")

	if "items displayed" not in msg:
		raise DataExportError(msg)

	_press_key("CtrlF8")        # open layout mgmt dialog
	_press_key("CtrlShiftF6")   # toggle technical names
	_press_key("Enter")         # Confirm Layout Changes
	_export_to_file(file)
	_press_key("F3")            # Load main mask
	data = _read_exported_data(file)

	try:
		os.remove(file)
	except (PermissionError, FileNotFoundError) as exc:
		log.error(exc)

	return data
