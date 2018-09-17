# encoding: utf-8

import logging
import os,sys

logging.basicConfig(level=logging.DEBUG, filename="adfc_rest2.log", filemode="w")
global logger
logger = logging.getLogger("adfc_rest2")
logger.info("cwd=%s", os.getcwd())
