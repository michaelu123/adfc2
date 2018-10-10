# encoding: utf-8

import logging
import os,sys

logging.basicConfig(level=logging.DEBUG, filename="adfc.log", filemode="w")
global logger
logger = logging.getLogger("adfc")
logger.info("cwd=%s", os.getcwd())
