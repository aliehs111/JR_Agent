from pywinauto import Application
import logging
import time

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

logger.debug("Connecting to FileMaker")
try:
    app = Application(backend="uia").connect(title_re=".*Pacific Solutionsv10.*", timeout=10)
    window = app.window(title_re=".*Pacific Solutionsv10.*")
    logger.debug("Waiting for window")
    window.wait("ready", timeout=10)
    logger.debug("Trying to find Purchase Orders button")
    button = window.child_window(auto_id="Group: 8598", control_type="Button")
    button.wait("enabled", timeout=5)
    button.draw_outline()  # Highlights the button for visual confirmation
    logger.debug("Clicking button")
    button.click_input()
    time.sleep(3)
    logger.debug("Done")
except Exception as e:
    logger.error(f"Error: {str(e)}")