from pywinauto import Application, Desktop
import time
import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

try:
    logger.debug("Listing all windows before connection")
    all_windows = [w.window_text() for w in Desktop(backend="uia").windows()]
    logger.debug(f"Windows: {all_windows}")
    logger.debug("Connecting to FileMaker")
    app = Application(backend="uia").connect(title="Pacific Solutionsv10 (CF-VM-App)", timeout=10)
    window = app.window(title="Pacific Solutionsv10 (CF-VM-App)")
    logger.debug("Waiting for window")
    window.wait("ready", timeout=10)
    logger.debug(f"Window found: {window.window_text()}")
    logger.debug("Dashboard controls:")
    window.print_control_identifiers()
    logger.debug("Clicking Purchase Orders button")
    window.child_window(title="Purchase Orders", control_type="Button").click()
    time.sleep(3)
    logger.debug("Post-navigation controls:")
    window.print_control_identifiers()
    logger.debug("Selecting Find from top bar")
    window.child_window(title="Find", control_type="Button").click()
    time.sleep(2)
    logger.debug("Typing query: >9/2/2025")
    window.type_keys(">9/2/2025{ENTER}")
    time.sleep(2)
    logger.debug("Clicking Excel button")
    window.child_window(title="Excel", control_type="Button").click()
    time.sleep(2)
    logger.debug("Listing all windows")
    all_windows = [w.window_text() for w in Desktop(backend="uia").windows()]
    logger.debug(f"Windows: {all_windows}")
    export_dialog = app.window(title_re="Export Records|Save As|Export.*|Save.*")
    logger.debug("Waiting for Export dialog")
    export_dialog.wait("ready", timeout=10)
    logger.debug("Export dialog found")
    export_dialog.print_control_identifiers()
    export_dialog.ComboBox.select("Comma-Separated Text (*.csv)")
    export_dialog.Edit.set_text(r"C:\Temp\export.csv")
    export_dialog.Save.click()
    time.sleep(2)
    logger.debug("Export completed")
except Exception as e:
    logger.error(f"Error: {str(e)}")