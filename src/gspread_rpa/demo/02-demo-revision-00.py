"""
High Level GoogleSheet API demo

apply some modification on a spreadsheet while listing the revisions.
save a specifique revision locally and re import it as a new spreadsheet using the 'id' returned by
file_upload

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

    for i in gs.revision_list_mtime():
        logger.info ("revision: {}".format(i))

    data = [[i*j for i in range(1,10)] for j in range(1,10)]
    logger.info ("data:\n {}".format("\n".join(["{}".format(i) for i in data])))

    csi = CellIndex(col=1, row=1)
    logger.info ("gs.update_cells(cells_index={}, values=data)".format(csi))
    gs.update_cells(cells_index=csi, values=data)
    gidx = None
    logger.info("gs.get_values()".format(gidx))
    result = gs.get_values(grid_index=gidx)
    logger.info("")
    for i in result:
        logger.info ("result: {}".format(i))
    for i in gs.revision_list_mtime():
        logger.info ("revision: {}".format(i))

    r = gs.revision_list_mtime()[-1]
    logger.info("save last revision: {}".format(r))
    with open (os.path.join(TMP, '{}-{}.ods'.format(os.path.basename(__file__), r.id)), 'wb') as fd:
        gs.file_export (fd, revision_id=r.id)

    logger.info("")

    data = [[i+j - 1 for i in range(1,10)] for j in range(1,10)]
    logger.info ("data:\n {}".format("\n".join(["{}".format(i) for i in data])))

    csi = CellIndex(col=1, row=1)
    logger.info ("gs.update_cells(cells_index={}, values=data)".format(csi))
    gs.update_cells(cells_index=csi, values=data)

    gidx = None
    logger.info("gs.get_values()".format(gidx))
    result = gs.get_values(grid_index=gidx)
    logger.info("")


    for i in gs.revision_list_mtime():
        logger.info ("revision: {}".format(i))

    logger.info("")

    up_file = None
    with open (os.path.join(TMP, '{}-{}.ods'.format(os.path.basename(__file__), r.id)), 'rb') as fd:
        logger.info ("upload: {}".format(fd.name))
        up_file = gs.file_upload (fd, title=os.path.basename(fd.name), return_object=False)
        logger.info ("upload: {} {}".format(fd.name, up_file))

    gsnew =  GoogleSheets()
    gsnew.open (key=up_file)
    logger.info ("gsnew: {}".format(gsnew.worksheets()))

    gidx=None
    gsnew.open(tab_name='mul')
    logger.info("gsnew.get_values()".format(gidx))
    result = gsnew.get_values(grid_index=gidx)
    logger.info("")
    for i in result:
        logger.info ("result: {}".format(i))
    for i in gsnew.revision_list_mtime():
        logger.info ("revision: {}".format(i))

    gs.remove_permission (email=demo_email)

    """
    cleanup
    """
    logger.info("")
    gsnew.delete_spreadsheet()
    gs.delete_spreadsheet()
