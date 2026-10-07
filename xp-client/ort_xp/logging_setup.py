# -*- coding: utf-8 -*-
"""日志。

Python 3.4 的 ``logging.FileHandler`` 不支持 ``encoding`` 参数（3.9 才有），
而日志里必然有简体中文，所以这里用 ``codecs.open`` 自己包一个 UTF-8 文件处理器；
控制台处理器走 ``compat.say``，在 cp950 控制台上不会因为编码崩掉。
"""

import codecs
import logging
import os

from . import compat

LOGGER_NAME = "ort_xp"

_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"


class Utf8FileHandler(logging.Handler):
    """UTF-8 文件日志（替代 3.4 上没有 encoding 参数的 FileHandler）。"""

    def __init__(self, path, mode="a"):
        logging.Handler.__init__(self)
        compat.ensure_dir(os.path.dirname(path))
        self.path = path
        self._stream = codecs.open(path, mode, "utf-8")

    def emit(self, record):
        try:
            self._stream.write(self.format(record) + os.linesep)
            self._stream.flush()
        except Exception:
            self.handleError(record)

    def close(self):
        try:
            self._stream.close()
        except Exception:
            pass
        logging.Handler.close(self)


class SafeConsoleHandler(logging.Handler):
    """控制台日志（编码安全）。"""

    def emit(self, record):
        try:
            compat.say(self.format(record))
        except Exception:
            pass


def setup_logging(app_directory, level=logging.INFO, console=True):
    """初始化日志，返回 logger。重复调用不会重复挂处理器。"""
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(level)
    logger.propagate = False
    if logger.handlers:
        return logger

    log_path = os.path.join(app_directory, "Logs", "ort_xp.log")
    try:
        file_handler = Utf8FileHandler(log_path, "a")
        file_handler.setFormatter(logging.Formatter(_FORMAT))
        logger.addHandler(file_handler)
    except Exception:
        # 日志目录不可写不应该影响程序启动
        pass

    if console:
        console_handler = SafeConsoleHandler()
        console_handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        logger.addHandler(console_handler)

    return logger


def get_logger():
    return logging.getLogger(LOGGER_NAME)
