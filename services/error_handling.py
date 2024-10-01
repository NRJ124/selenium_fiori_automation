import logging

logger = logging.getLogger(__name__)

def handle_webdriver_error(e):
    logger.error(f"WebDriver error occurred: {e}")
    # Optionally retry or handle it in a different way
