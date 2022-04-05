"""
High Level GoogleSheet API demo
demo - how to get multiples cells format as a list of list
     - Cellformat copy to aplly the format only once
       (not running gs.apply_cells_user_format() just after each changes)
"""

import os, sys
import logging
from logging.handlers import TimedRotatingFileHandler
from tempfile import gettempdir

if os.path.exists(
        os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'), 'gspread_rpa')):
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
from gspread_rpa import CellIndex, GridIndex, GoogleSheets, CellFormat

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

    # reset formatting
    gs.prepare_cells_user_format (GridIndex(1,20,1,20), CellFormat())
    gs.apply_cells_user_format()
    cf_mat = gs.get_cells_user_format (GridIndex(1,20,1,20))
    logger.info ("empty cf_mat: {}".format(cf_mat))
    assert cf_mat == []

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

    gs.prepare_cells_user_format (GridIndex(1,1,1,1), cf01)

    cf02 = CellFormat (cf01)
    cf02.background_color ("green")

    gs.prepare_cells_user_format (GridIndex(9,1,9,1), cf02)

    cf03 = CellFormat (cf01)
    cf03.background_color ("yellow")

    gs.prepare_cells_user_format (GridIndex(9,9,9,9), cf03)

    cf04 = CellFormat (cf01)
    cf04.background_color ("red")

    gs.prepare_cells_user_format (GridIndex(1,9,1,9), cf04)

    gs.apply_cells_user_format()

    csi = GridIndex(start_col=1, start_row=1, end_col=9, end_row=9)
    cf_mat = gs.get_cells_user_format (csi)
    for i in cf_mat:
        logger.info ([j.__class__.__name__ for j in i])
    logger.info ("")
    for i in cf_mat:
        logger.info ([j.background_color_name if isinstance(j, CellFormat) else '' for j in i ])

    assert cf_mat[0][0].background_color_name == "blue"
    assert cf_mat[0][8].background_color_name == "green"
    assert cf_mat[8][0].background_color_name == "red"
    assert cf_mat[8][8].background_color_name == "yellow"

    gs.remove_permission (email=demo_email)

    """
    cleanup
    """
    try:
        gs.delete_spreadsheet()
    except Exception as e:
        logger.warning (e)
