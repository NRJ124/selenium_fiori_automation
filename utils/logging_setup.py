import logging
def setup_logger(log_file=r"C:\Users\enesine\PycharmProjects\selenium_fiori_automation\logs\application.log"):
    logging.basicConfig(filename=log_file,
                        level=logging.INFO,
                        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
                        datefmt="%Y-%m-%d %H:%M:%S")
    logger = logging.getLogger()
    return logger
