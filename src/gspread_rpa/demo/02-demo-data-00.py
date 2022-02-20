"""
High Level GoogleSheet API demo
"""

import os, sys
import logging
from logging.handlers import TimedRotatingFileHandler

sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
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

    gs = GoogleSheets()

    """
    open or create a new spreadsheet
    """

    try:
        gs.open (title=title)
    except gs.NotFound as e:
        gs.create (title=title)

    gs.give_permission (email=demo_email, perm_type='user', role='writer')

    for p in gs.list_permission():
        logger.info ("permission: {}".format([p[i] for i in ['type', 'role', 'emailAddress']]))

    """
    add worksheets
    """
    gs.add_worksheet(title='mul', rows=10 ,  cols=10,   tab_position=0)
    logger.info (gs.worksheet_cursor)

    for i in range(1,10):
        for j in range (1,4):
            gs.update_cell (col=i, row=j, value=i * j)

    for k in [1, 2]:
        logger.info ("get_values_row(1): {}".format (gs.get_values_row(1)))
        logger.info ("get_values_row(2): {}".format (gs.get_values_row(2)))
        logger.info ("get_values_row(3): {}".format (gs.get_values_row(3)))
        for i in range(1,10):
            logger.info ("get_values_col({}): {}".format (i, gs.get_values_col(i)))
        if k == 2: break
        idx = [
            GridIndex(start_col=5, start_row=1, end_col=6, end_row=1),
            CellIndex(col=8, row=2), GridIndex(start_col=3, start_row=2, end_col=4, end_row=3)
        ]
        logger.info ("clear ({})".format (idx))
        gs.clear (idx)
        exp = [
            ['1', '2', '3', '4' , ''   , ''  , '7'  , '8' , '9' ],
            ['2', '4', '', ''   , '10' , '12', '14' , ''  , '18'],
            ['3', '6', '', ''   , '15' , '18', '21' , '24', '27']]

        res = gs.get_values(idx)
        assert gs.get_values(idx) == exp

    gs.remove_permission (email=demo_email)

    """
    cleanup
    """
    gs.delete_spreadsheet()
