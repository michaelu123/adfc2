# encoding: utf-8

import logging
import os

logging.basicConfig(level=logging.DEBUG, filename="adfc_rest1.log", filemode="w")
global logger
logger = logging.getLogger("adfc-rest1")
logger.info("cwd=%s", os.getcwd())
