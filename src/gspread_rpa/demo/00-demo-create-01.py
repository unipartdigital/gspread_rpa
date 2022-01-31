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

    runs = [i for i in range(1,4)]
    for run in runs:
        logger.info ("run {}/{}".format(run, runs[-1]))
        gs = GoogleSheets()

        """
        open or create a new spreadsheet
        """

        try:
            gs.open (title=title)
        except gs.NotFound as e:
            if run == runs[0]:
                gs.create (title=title)
            else:
                raise ProgramError

        gs.give_permission (email=demo_email, perm_type='user', role='writer')

        for p in gs.list_permission():
            logger.info ("permission: {}".format([p[i] for i in ['type', 'role', 'emailAddress']]))

        """
        add worksheets
        """
        gs.add_worksheet(title='ws0', rows=10 ,  cols=10,   tab_position=0)

        gs.add_worksheet(title='ws1', rows=100,  cols=100,  tab_position=1)
        """ by default doesn't raise if the worksheet  already exist unless raise_if_exists """
        gs.add_worksheet(title='ws1', rows=100,  cols=100,  tab_position=1)

        try:
            gs.add_worksheet(title='ws1', rows=100,  cols=100,  tab_position=1, raise_if_exists=True)
        except gs.AlreadyExists as e:
            logger.warning ("{} already exits".format('ws1'))

        gs.add_worksheet(title='ws2', rows=1000, cols=1000, tab_position=2)

        wslist = gs.worksheets()
        wslist.reverse()
        gs.reorder_worksheets(wslist)

        gs.remove_permission (email=demo_email)

        """
        cleanup
        """
        for ws in ['ws0', 'ws1', 'ws2']:
            """ last opened worksheet become the active worksheet """
            gs.open (tab_name=ws)
            gs.delete_worksheet()
        if run == runs[-1]:
            gs.delete_spreadsheet()
