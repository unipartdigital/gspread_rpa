"""
High Level GoogleSheet API demo
"""

import os, sys
import logging
from logging.handlers import TimedRotatingFileHandler
from tempfile import gettempdir

sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
from gspread_rpa import CellIndex, GridIndex, GoogleSheets

TMP = gettempdir()
assert os.path.isdir (TMP)

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
    gs.add_worksheet(title='mul', rows=20 ,  cols=20,   tab_position=0)
    logger.info (gs.worksheet_cursor)


    data = [[i*j for i in range(1,10)] for j in range(1,10)]
    logger.info ("data:\n {}".format("\n".join(["{}".format(i) for i in data])))

    csi = CellIndex(col=1, row=1)
    logger.info ("gs.update_cells(cells_index={}, values=data)".format(csi))
    gs.update_cells(cells_index=csi, values=data)

    cf01 = gs.get_cell_user_format (csi)
    logger.info ("cf01.o2dict(): {}".format(cf01.o2dict()))

    cf01.background_color ("blue")
    cf01.text.foreground_color("green")
    cf01.text.bold(True)
    cf01.text.font_family("Arial")
    cf01.text.font_size(12)
    cf01.text.italic(False)

    gs.prepare_cells_user_format (GridIndex(1,5,3,9), cf01)
    gs.apply_cells_user_format()

    csi = CellIndex(col=1, row=5)
    cf02 = gs.get_cell_user_format (csi)
    logger.info ("cf02: {}".format(cf02.o2dict()))

    assert cf01.background_color_name == cf02.background_color_name

    gs.remove_permission (email=demo_email)

    """
    cleanup
    """
    try:
        gs.delete_spreadsheet()
    except Exception as e:
        logger.warning (e)
