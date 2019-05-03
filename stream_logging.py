# ----------------------------------------------------------------------------#
# (C) British Crown Copyright 2019 Met Office.                                #
# Author: Steve Wardle                                                        #
#                                                                             #
# This file is part of OWA Checker.                                           #
# OWA Checker is free software: you can redistribute it and/or modify it      #
# under the terms of the Modified BSD License, as published by the            #
# Open Source Initiative.                                                     #
#                                                                             #
# OWA Checker is distributed in the hope that it will be useful,              #
# but WITHOUT ANY WARRANTY; without even the implied warranty of              #
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the               #
# Modified BSD License for more details.                                      #
#                                                                             #
# You should have received a copy of the Modified BSD License                 #
# along with OWA Checker...                                                   #
# If not, see <http://opensource.org/licenses/BSD-3-Clause>                   #
# ----------------------------------------------------------------------------#
import os
import sys
import logging
import logging.handlers

# Each log file produced will not exceed this size
MAX_BYTES_PER_FILE = 1024 * 16


class LogStream(object):
    """
    An object which behaves like a file stream, but writes to a logger

    """
    def __init__(self, logger, level=logging.INFO):
        self.logger = logger
        self.level = level

    def write(self, buf):
        for line in buf.rstrip().splitlines():
            self.logger.log(self.level, line.rstrip())

    def flush(self):
        pass


def create_logger(filename):
    """
    Initialise the logfile and logger
    """
    if os.path.exists(filename):
        os.remove(filename)
    hdl = logging.handlers.RotatingFileHandler(
        filename, maxBytes=MAX_BYTES_PER_FILE)
    formatter = logging.Formatter(
        '%(asctime)s, %(levelname)s %(message)s', "%Y-%m-%d %H:%M:%S")
    hdl.setFormatter(formatter)
    logger = logging.getLogger("owa_checker")
    logger.addHandler(hdl)
    logger.setLevel(logging.DEBUG)

    # Attach both stderr and stdout to the logger
    sys.stderr = LogStream(logger, logging.ERROR)
    sys.stdout = LogStream(logger, logging.INFO)

    return logger
