from pywinauto import Application
import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

try:
    app = Application(backend="uia").connect(title="Pacific Solutionsv10 (CF-VM-App)", timeout=10)
    window = app.window(title="Pacific Solutionsv10 (CF-VM-App)")
    window.wait("ready", timeout=10)
    logger.debug("Dashboard controls:")
    window.print_control_identifiers()
except Exception as e:
    logger.error(f"Error: {str(e)}")