from pywinauto import Application, Desktop
from pywinauto.findwindows import ElementNotFoundError
import time
import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def try_connect(backend, title_re, timeout=10):
    try:
        app = Application(backend=backend).connect(title_re=title_re, timeout=timeout)
        window = app.window(title_re=title_re)
        window.wait("ready", timeout=timeout)
        logger.debug(f"Connected with {backend} backend: {window.window_text()}")
        return app, window, backend
    except Exception as e:
        logger.error(f"Failed to connect with {backend} backend: {str(e)}")
        return None, None, backend

try:
    # Step 1: List all windows
    logger.debug("Listing all windows before connection")
    all_windows = [w.window_text() for w in Desktop(backend="uia").windows()]
    logger.debug(f"Windows: {all_windows}")

    # Step 2: Connect to FileMaker
    logger.debug("Connecting to FileMaker")
    app, window, backend_used = try_connect("uia", ".*Pacific Solutionsv10.*")
    if not app:
        logger.debug("Trying win32 backend")
        app, window, backend_used = try_connect("win32", ".*Pacific Solutionsv10.*")
    if not app:
        logger.error("Failed to connect with uia or win32 backend")
        raise Exception("Failed to connect with uia or win32 backend")

    # Step 3: Save dashboard controls
    logger.debug(f"Dashboard controls ({backend_used}):")
    window.print_control_identifiers(filename=f"dashboard_controls_{backend_used}.txt")

    # Step 4: Test Purchase Orders button
    auto_ids_to_test = [
        "Group: 48366", "Group: 48372", "Group: 48378", "Group: 48384",  # From dashboard_controls_uia.txt
        "Group: 48390", "Group: 48396", "Group: 48402", "Group: 48408"  # Top buttons
    ]
    for auto_id in auto_ids_to_test:
        logger.debug(f"Testing button with auto_id={auto_id} ({backend_used})")
        try:
            button = window.child_window(auto_id=auto_id, control_type="Button")
            button.wait("enabled", timeout=5)
            button.draw_outline()
            logger.debug(f"Clicking button {auto_id}")
            button.click_input()
            time.sleep(3)
            logger.debug(f"Controls after clicking {auto_id}:")
            window.print_control_identifiers(filename=f"controls_after_{auto_id.replace(':', '_')}_{backend_used}.txt")
            logger.debug("Resetting to dashboard (manually ensure FileMaker is back on dashboard)")
            input("Press Enter after resetting FileMaker to dashboard...")
            app, window, _ = try_connect(backend_used, ".*Pacific Solutionsv10.*")
            if not app:
                logger.error(f"Failed to reconnect with {backend_used} after clicking {auto_id}")
                continue
        except ElementNotFoundError as e:
            logger.error(f"Button {auto_id} not found ({backend_used}): {str(e)}")
            continue
        except Exception as e:
            logger.error(f"Unexpected error for {auto_id} ({backend_used}): {str(e)}")
            continue

    # Step 5: Try coordinate-based clicks
    logger.debug(f"Trying coordinate-based clicks ({backend_used})")
    coords_to_test = [
        (1202, 585),  # Center of Group: 48366 (L1149, T568, R1256, B603)
        (1202, 700),  # Center of Group: 48372 (L1149, T683, R1256, B718)
        (1202, 815),  # Center of Group: 48378 (L1149, T798, R1256, B833)
        (1202, 875)   # Center of Group: 48384 (L1149, T858, R1256, B893)
    ]
    for coords in coords_to_test:
        logger.debug(f"Testing click at coordinates {coords} ({backend_used})")
        try:
            window.click_input(coords=coords)
            time.sleep(3)
            logger.debug(f"Controls after clicking at {coords}:")
            window.print_control_identifiers(filename=f"controls_after_coords_{coords[0]}_{coords[1]}_{backend_used}.txt")
            logger.debug("Resetting to dashboard (manually ensure FileMaker is back on dashboard)")
            input("Press Enter after resetting FileMaker to dashboard...")
            app, window, _ = try_connect(backend_used, ".*Pacific Solutionsv10.*")
            if not app:
                logger.error(f"Failed to reconnect with {backend_used} after clicking at {coords}")
                continue
        except Exception as e:
            logger.error(f"Error with click at {coords} ({backend_used}): {str(e)}")
            continue

    # Step 6: Try title-based Purchase Orders button
    logger.debug(f"Trying title-based Purchase Orders button ({backend_used})")
    try:
        button = window.child_window(title="Purchase Orders", control_type="Button")
        button.wait("enabled", timeout=5)
        button.draw_outline()
        logger.debug("Clicking Purchase Orders button")
        button.click_input()
        time.sleep(3)
        logger.debug(f"Controls after clicking Purchase Orders ({backend_used}):")
        window.print_control_identifiers(filename=f"controls_after_purchase_orders_{backend_used}.txt")
        logger.debug("Resetting to dashboard (manually ensure FileMaker is back on dashboard)")
        input("Press Enter after resetting FileMaker to dashboard...")
        app, window, _ = try_connect(backend_used, ".*Pacific Solutionsv10.*")
        if not app:
            logger.error(f"Failed to reconnect with {backend_used} after title-based click")
    except ElementNotFoundError as e:
        logger.error(f"Purchase Orders button not found by title ({backend_used}): {str(e)}")
    except Exception as e:
        logger.error(f"Unexpected error for title-based Purchase Orders ({backend_used}): {str(e)}")

    # Step 7: Manual navigation to confirm Purchase Orders controls
    logger.debug("Please manually click Purchase Orders and press Enter")
    input("Press Enter after navigating to Purchase Orders...")
    logger.debug(f"Trying to reconnect to FileMaker after manual navigation ({backend_used})")
    app, window, _ = try_connect(backend_used, ".*Pacific Solutionsv10.*|.*Purchase Orders.*")
    if not app:
        all_windows = [w.window_text() for w in Desktop(backend=backend_used).windows()]
        logger.debug(f"Windows after manual navigation: {all_windows}")
        app, window, _ = try_connect(backend_used, ".*")
        if not app:
            logger.error(f"Failed to reconnect with {backend_used} after manual navigation")
            raise Exception(f"Failed to reconnect with {backend_used} after manual navigation")
    logger.debug(f"Post-Purchase Orders controls ({backend_used}):")
    window.print_control_identifiers(filename=f"post_purchase_orders_{backend_used}.txt")

    # Step 8: Test Find button
    logger.debug(f"Selecting Find button ({backend_used})")
    try:
        find_button = window.child_window(title="Find", control_type="Button")
        find_button.wait("enabled", timeout=5)
        find_button.draw_outline()
        find_button.click_input()
        time.sleep(2)
        logger.debug(f"Post-Find controls ({backend_used}):")
        window.print_control_identifiers(filename=f"post_find_controls_{backend_used}.txt")
    except ElementNotFoundError as e:
        logger.error(f"Find button not found ({backend_used}): {str(e)}")
    except Exception as e:
        logger.error(f"Unexpected error for Find button ({backend_used}): {str(e)}")

    # Step 9: Type query
    logger.debug("Typing query: >9/2/2025")
    try:
        window.type_keys(">9/2/2025{ENTER}")
        time.sleep(2)
    except Exception as e:
        logger.error(f"Error typing query: {str(e)}")

    # Step 10: Test Excel button
    logger.debug(f"Clicking Excel button ({backend_used})")
    try:
        excel_button = window.child_window(title="Excel", control_type="Button")
        excel_button.wait("enabled", timeout=5)
        excel_button.draw_outline()
        excel_button.click_input()
        time.sleep(2)
        logger.debug(f"Post-Excel controls ({backend_used}):")
        window.print_control_identifiers(filename=f"post_excel_controls_{backend_used}.txt")
    except ElementNotFoundError as e:
        logger.error(f"Excel button not found ({backend_used}): {str(e)}")
    except Exception as e:
        logger.error(f"Unexpected error for Excel button ({backend_used}): {str(e)}")

    # Step 11: Handle Export dialog
    logger.debug("Listing all windows")
    all_windows = [w.window_text() for w in Desktop(backend=backend_used).windows()]
    logger.debug(f"Windows: {all_windows}")
    try:
        export_dialog = app.window(title_re="Export Records|Save As|Export.*|Save.*")
        logger.debug("Waiting for Export dialog")
        export_dialog.wait("ready", timeout=10)
        logger.debug(f"Export dialog controls ({backend_used}):")
        export_dialog.print_control_identifiers(filename=f"export_dialog_controls_{backend_used}.txt")
        logger.debug("Selecting CSV format")
        try:
            export_dialog.ComboBox.select("Comma-Separated Text (*.csv)")
            export_dialog.Edit.set_text(r"C:\Temp\export.csv")
            export_dialog.child_window(title="Save", control_type="Button").click()
            time.sleep(2)
            logger.debug("Export completed")
        except ElementNotFoundError as e:
            logger.error(f"Export dialog elements not found ({backend_used}): {str(e)}")
        except Exception as e:
            logger.error(f"Unexpected error for Export dialog ({backend_used}): {str(e)}")
    except Exception as e:
        logger.error(f"Error accessing Export dialog: {str(e)}")
except Exception as e:
    logger.error(f"Error: {str(e)}")