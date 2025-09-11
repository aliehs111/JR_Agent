from pywinauto import Application
from pywinauto.keyboard import send_keys
import time
import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

try:
    logger.debug("Connecting to FileMaker")
    app = Application(backend="uia").connect(title_re=".*Pacific Solutionsv10.*", timeout=10)
    window = app.window(title_re=".*Pacific Solutionsv10.*")
    logger.debug("Waiting for window")
    window.wait("ready", timeout=10)

    logger.debug("Navigating to Purchase Orders via menu")
    window.set_focus()
    send_keys("%v")  # Alt+V
    time.sleep(1)
    send_keys("{DOWN}")  # Adjust number of DOWN presses if needed
    send_keys("{ENTER}")
    time.sleep(3)
    logger.debug("Post-Purchase Orders controls:")
    window.print_control_identifiers(filename="post_purchase_orders_uia.txt")

    logger.debug("Entering Find mode")
    send_keys("^f")  # Ctrl+F
    time.sleep(2)
    logger.debug("Typing query: >9/2/2025")
    send_keys(">9/2/2025{ENTER}")
    time.sleep(2)
    logger.debug("Post-Find controls:")
    window.print_control_identifiers(filename="post_find_controls_uia.txt")

    logger.debug("Listing all windows before export")
    all_windows = [w.window_text() for w in Desktop(backend="uia").windows()]
    logger.debug(f"Windows: {all_windows}")

    logger.debug("Trying export via File menu")
    try:
        send_keys("%f")  # Alt+F
        time.sleep(1)
        send_keys("e")   # E for Export
        time.sleep(2)
    except Exception as e:
        logger.error(f"File menu export failed: {str(e)}")

    logger.debug("Trying export via Records menu")
    try:
        send_keys("%r")  # Alt+R
        time.sleep(1)
        send_keys("e")   # E for Export
        time.sleep(2)
    except Exception as e:
        logger.error(f"Records menu export failed: {str(e)}")

    logger.debug("Listing all windows after export attempts")
    all_windows = [w.window_text() for w in Desktop(backend="uia").windows()]
    logger.debug(f"Windows: {all_windows}")

    logger.debug("Opening Export dialog")
    try:
        export_dialog = app.window(title_re="Export Records|Save As|Export.*|Save.*")
        logger.debug("Waiting for Export dialog")
        export_dialog.wait("ready", timeout=15)
        logger.debug("Export dialog controls:")
        export_dialog.print_control_identifiers(filename="export_dialog_controls_uia.txt")
        logger.debug("Selecting CSV format")
        export_dialog.ComboBox.select("Comma-Separated Text (*.csv)")
        export_dialog.Edit.set_text(r"C:\Temp\export.csv")
        export_dialog.child_window(title="Save", control_type="Button").click()
        time.sleep(2)
        logger.debug("Export completed")
    except Exception as e:
        logger.error(f"Error accessing Export dialog: {str(e)}")
except Exception as e:
    logger.error(f"Error: {str(e)}")