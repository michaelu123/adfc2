# encoding: utf-8

import logging
import os,sys

try:
    lvl = os.environ["DEBUG"]
    lvl = logging.DEBUG
except:
    lvl = logging.ERROR
logging.basicConfig(level=lvl, filename="adfc.log", filemode="w", datefmt="%d.%m %H:%M:%S",
    style="{", format="{asctime} {levelname:5} {filename}:{funcName}:{lineno} {message}")
global logger
logger = logging.getLogger("adfc")
logger.info("cwd=%s", os.getcwd())
