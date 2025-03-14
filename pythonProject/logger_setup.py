import logging
from logging import Logger
from typing import Tuple


class MatplotlibFilter(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        # Игнорируем сообщения от matplotlib.category
        return 'matplotlib.category' not in record.name


def setup_logger() -> tuple[Logger, Logger]:
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    # # Настраиваем базовую конфигурацию логирования для ГК
    # logging.basicConfig(level=logging.INFO,
    #                     format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    #                     datefmt='%a, %d %b %Y %H:%M:%S',
    #                     filename='log_GK.txt',
    #                     filemode='a+'
    #                     )

    logger_FDA = logging.getLogger('ФДА Росавтодор')
    if logger_FDA.hasHandlers():
        logger_FDA.handlers.clear()

    file_handler_FDA = logging.FileHandler('log_FDA.txt', mode='a+')
    formatter_FDA = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                                   datefmt='%a, %d %b %Y %H:%M:%S')
    logger_FDA.setLevel(logging.INFO)
    file_handler_FDA.setFormatter(formatter_FDA)
    logger_FDA.addHandler(file_handler_FDA)

    # # Настраиваем базовую конфигурацию логирования для ФДА
    # logging.basicConfig(level=logging.INFO,
    #                     format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    #                     datefmt='%a, %d %b %Y %H:%M:%S',
    #                     filename='log_FDA.txt',
    #                     filemode='a+'
    #                     )

    # Настраиваем второй логгер для log_FDA.txt
    logger_GK = logging.getLogger('ГК Автодор')
    if logger_GK.hasHandlers():
        logger_GK.handlers.clear()
    file_handler_GK = logging.FileHandler('log_GK.txt', mode='a+')
    formatter_GK = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                                    datefmt='%a, %d %b %Y %H:%M:%S')
    logger_GK.setLevel(logging.INFO)
    file_handler_GK.setFormatter(formatter_GK)
    logger_GK.addHandler(file_handler_GK)

    # logger_FDA = logging.getLogger('ФДА Росавтодор')
    # logger_GK = logging.getLogger('ГК Автодор')

    mlogger = logging.getLogger('matplotlib')
    mlogger.setLevel(logging.WARNING)

    matplotlib_logger = logging.getLogger("matplotlib.category")
    matplotlib_logger.addFilter(MatplotlibFilter())

    return logger_FDA, logger_GK


# def setup_logger():
#     for handler in logging.root.handlers[:]:
#         logging.root.removeHandler(handler)
#
#     logging.basicConfig(level=logging.INFO,
#                         format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
#                         datefmt='%a, %d %b %Y %H:%M:%S',
#                         filename='log_GK.txt',
#                         filemode='a+'
#                         )
#
#     logger = logging.getLogger('my_logger')
#     return logger