# encoding: utf-8

import logging
import os

"""
see https://stackoverflow.com/questions/10706547/add-encoding-parameter-to-logging-basicconfig"""
global logger
logger = logging.getLogger("adfc")
try:
    lvl = os.environ["DEBUG"]
    lvl = logging.DEBUG
except:
    lvl = logging.ERROR
logger.setLevel(lvl) # or whatever
handler = logging.FileHandler('adfc.log', 'w', 'utf-8') # need utf encoding
formatter = logging.Formatter(style="$", datefmt="%d.%m %H:%M:%S", fmt="${asctime} ${levelname} ${filename}:${funcName}:${lineno} ${message}")
handler.setFormatter(formatter) # Pass handler as a parameter, not assign
logger.addHandler(handler)
logger.info("cwd=%s", os.getcwd())
