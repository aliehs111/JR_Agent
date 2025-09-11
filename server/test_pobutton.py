from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
import time
import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

try:
    logger.debug("Connecting to FileMaker Navigator")
    app = Application(backend="uia").connect(title_re=".*Pacific Solutionsv10.*", timeout=10)
    filemaker_window = app.window(title_re=".*Pacific Solutionsv10.*")
    logger.debug("Waiting for Navigator window")
    filemaker_window.wait("ready", timeout=10)

    logger.debug("Listing all windows before Purchase Orders click")
    all_windows = [w.window_text() for w in Desktop(backend="uia").windows()]
    logger.debug(f"Windows: {all_windows}")

    logger.debug("Clicking Purchase Orders button at (1528, 726)")
    filemaker_window.set_focus()
    filemaker_window.click_input(coords=(1528, 726))
    time.sleep(5)

    logger.debug("Switching to POv10 window")
    filemaker_window = app.window(title_re="POv10.*")
    filemaker_window.wait("ready", timeout=15)
    logger.debug("Reached POv10 window")

    logger.debug("Clicking first Find button (Name: Find 3 of 11, ControlType: RadioButton)")
    find_button = filemaker_window.child_window(title="Find 3 of 11", control_type="RadioButton")
    find_button.click_input()
    time.sleep(3)

    logger.debug("Listing all windows after Find click")
    all_windows = [w.window_text() for w in Desktop(backend="uia").windows()]
    logger.debug(f"Windows: {all_windows}")

    logger.debug("Clicking Order Date field at (1259, 571, ControlType: Edit)")
    try:
        order_date_field = filemaker_window.child_window(control_type="Edit", found_index=1)
        order_date_field.click_input()
        time.sleep(1)
        logger.debug("Typing query: >9/2/2025")
        send_keys(">9/2/2025")
        time.sleep(1)
    except Exception as e:
        logger.error(f"Error clicking Order Date field by control: {str(e)}")
        logger.debug("Trying coordinates (1259, 571)")
        filemaker_window.click_input(coords=(1259, 571))
        time.sleep(1)
        logger.debug("Typing query: >9/2/2025")
        send_keys(">9/2/2025")
        time.sleep(1)

    logger.debug("Clicking second Find button (Name: Generic IView, ControlType: Pane)")
    try:
        query_find_button = filemaker_window.child_window(title="Generic IView", control_type="Pane")
        query_find_button.click_input()
        time.sleep(3)
    except Exception as e:
        logger.error(f"Error clicking second Find button by control: {str(e)}")
        logger.debug("Trying coordinates (1355, 946)")
        filemaker_window.click_input(coords=(1355, 946))
        time.sleep(3)

    logger.debug("Listing all windows after query Find click")
    all_windows = [w.window_text() for w in Desktop(backend="uia").windows()]
    logger.debug(f"Windows: {all_windows}")

    logger.debug("Clicking Excel button (ControlType: Button near Excel)")
    try:
        excel_button = filemaker_window.child_window(auto_id="Group: 244677", control_type="Button")
        excel_button.click_input()
        time.sleep(5)
    except Exception as e:
        logger.error(f"Error clicking Excel button by control (Group: 244677): {str(e)}")
        logger.debug("Trying alternative button (Group: 244692)")
        try:
            excel_button_alt = filemaker_window.child_window(auto_id="Group: 244692", control_type="Button")
            excel_button_alt.click_input()
            time.sleep(5)
        except Exception as e:
            logger.error(f"Error clicking alternative Excel button (Group: 244692): {str(e)}")
            logger.debug("Trying coordinates (407, 55)")
            filemaker_window.click_input(coords=(407, 55))
            time.sleep(5)

    logger.debug("Listing all windows after Excel click")
    all_windows = [w.window_text() for w in Desktop(backend="uia").windows()]
    logger.debug(f"Windows: {all_windows}")

    logger.debug("Clicking PO List button in Export dialog (Name: PO List, ControlType: Button, AutomationId: IDC_SHOW_CUSTOM_MESSAGE_BUTTON3)")
    try:
        export_dialog = app.window(title="Export", control_type="Window")
        export_dialog.wait("ready", timeout=15)
        po_list_button = export_dialog.child_window(title="PO List", auto_id="IDC_SHOW_CUSTOM_MESSAGE_BUTTON3", control_type="Button")
        po_list_button.click_input()
        time.sleep(3)
    except Exception as e:
        logger.error(f"Error clicking PO List button by control: {str(e)}")
        logger.debug("Trying coordinates (1241, 741) for PO List")
        export_dialog = app.window(title="Export", control_type="Window")
        export_dialog.click_input(coords=(1241, 741))
        time.sleep(3)
        logger.debug("Trying alternative coordinates (1313, 741) for PO Items")
        export_dialog.click_input(coords=(1313, 741))
        time.sleep(3)

    logger.debug("Listing all windows after PO List click")
    all_windows = [w.window_text() for w in Desktop(backend="uia").windows()]
    logger.debug(f"Windows: {all_windows}")

    logger.debug("Saving controls after PO List click")
    try:
        save_dialog = app.window(title_re="Save As")
        save_dialog.wait("ready", timeout=10)
        save_dialog.print_control_identifiers(filename="post_po_list_button_uia.txt")
        logger.debug("Success: Reached Save As dialog")
    except Exception as e:
        logger.error(f"Error accessing Save As dialog: {str(e)}")
        logger.debug("Saving controls from current window")
        filemaker_window.print_control_identifiers(filename="post_po_list_button_uia.txt")
except Exception as e:
    logger.error(f"Error: {str(e)}")