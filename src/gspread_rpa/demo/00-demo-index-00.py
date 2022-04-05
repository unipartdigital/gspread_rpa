"""
High Level GoogleSheet API demo
"""

import os, sys
import logging
from logging.handlers import TimedRotatingFileHandler

if os.path.exists(
        os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'), 'gspread_rpa')):
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
from gspread_rpa import CellIndex, GridIndex, GoogleSheets

demo_email = 'test@example.org'

logger = logging.getLogger(os.path.basename(__file__))

if __name__ == "__main__":

    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    logger_handlers = [logging.StreamHandler()]
    logger_handlers.append (TimedRotatingFileHandler(filename="{}.log".format(os.path.basename(__file__)),
                                                     when='D', # 'H' Hours 'D' Days
                                                     interval=1, backupCount=0,
                                                     encoding=None, utc=True))

    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s %(levelname)-8s %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S',
                        handlers=logger_handlers)

    title=os.path.basename(__file__.replace (".py", ''))


    try:
        logger.info ("")
        ci = CellIndex(col=0, row=0)
        logger.info ("expect error {}".format(ci))
        logger.info ("ci.to_a1: {}".format(ci.to_a1()))
        sys.exit(1)
    except Exception as e:
        logger.error ("ci.to_a1: {}".format(e))

    logger.info ("")
    ci = CellIndex(col=1, row=1)
    logger.info (ci)
    logger.info ("ci.to_a1: {}".format(ci.to_a1()))
    assert ci.to_a1() == "A1"
    logger.info("CellIndex().from_a1 ('A1') {}".format(CellIndex().from_a1 ("A1")))
    assert ci == CellIndex().from_a1 ("A1")

    logger.info ("")
    ci = CellIndex(col=100, row=200)
    logger.info (ci)
    logger.info ("ci.to_a1: {}".format(ci.to_a1()))
    assert ci.to_a1() == "CV200"

    logger.info ("")
    gi1 = GridIndex(start_col=1, start_row=1, end_col=1, end_row=1)
    gi2 = GridIndex(start_col=1, start_row=1, end_col=1, end_row=1)
    gi3 = GridIndex(start_col=2, start_row=1, end_col=1, end_row=1)
    gi4 = GridIndex(start_col=1, start_row=2, end_col=1, end_row=1)
    gi5 = GridIndex(start_col=1, start_row=1, end_col=2, end_row=1)
    gi6 = GridIndex(start_col=1, start_row=1, end_col=1, end_row=2)
    logger.info ("assert {} == {}".format(gi1, gi2))
    assert gi1 == gi2
    logger.info ("assert {} != {}".format(gi1, gi3))
    assert gi1 != gi3
    logger.info ("assert {} != {}".format(gi1, gi4))
    assert gi1 != gi4
    logger.info ("assert {} != {}".format(gi1, gi5))
    assert gi1 != gi5
    logger.info ("assert {} != {}".format(gi1, gi6))
    assert gi1 != gi6

    logger.info ("assert gi1.start.to_a1() == 'A1'")
    assert gi1.start.to_a1() == "A1"
    logger.info ("assert gi1.end.to_a1() == 'A1'")
    assert gi1.end.to_a1() == "A1"
    logger.info ("assert gi6.start.to_a1() == 'A1'")
    assert gi6.start.to_a1() == "A1"
    logger.info ("assert gi6.end.to_a1() == 'A2'")
    assert gi6.end.to_a1() == "A2"
