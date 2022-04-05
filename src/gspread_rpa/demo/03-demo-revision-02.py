"""
High Level GoogleSheet API demo

apply some modification on a spreadsheet while listing the revisions.
save a specifique revision locally and re import it in the same spreadsheet
using     gs.overwrite(gsnew, delete=True)

spreadsheet ID and worksheets name should stay stable.

"""

import os, sys
import logging
from logging.handlers import TimedRotatingFileHandler
from tempfile import gettempdir

if os.path.exists(
        os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'), 'gspread_rpa')):
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
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

    gs.delete_spreadsheet()
    try:
        gs.open (title=title)
    except gs.NotFound as e:
        gs.create (title=title)

    id_ori = gs.spreadsheet_cursor.id

    gs.give_permission (email=demo_email, perm_type='user', role='writer')

    for p in gs.list_permission():
        logger.info ("permission: {}".format([p[i] for i in ['type', 'role', 'emailAddress']]))

    """
    add worksheets
    """
    gs.add_worksheet(title='mul', rows=20 ,  cols=20,   tab_position=0)
    logger.info (gs.worksheet_cursor)

    ws_names_ori = gs.worksheets(only='name')
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
    for i in result:
        logger.info ("result: {}".format(i))

    for i in gs.revision_list_mtime():
        logger.info ("revision: {}".format(i))

    logger.info("")

    gsnew = None
    with open (os.path.join(TMP, '{}-{}.ods'.format(os.path.basename(__file__), r.id)), 'rb') as fd:
        logger.info ("upload: {}".format(fd.name))
        gsnew = gs.file_upload (fd, title=os.path.basename(fd.name))
        logger.info ("upload: {} {}".format(fd.name, gsnew))

    logger.info ("worksheets: {}".format(gs.worksheets()))
    logger.info ("gs.overwrite({}, delete=True)".format (gsnew))
    gs.overwrite(gsnew, delete=True)
    logger.info ("worksheets: {}".format(gs.worksheets()))

    gidx=None
    gs.open(tab_name='mul')
    logger.info("gs.get_values()".format(gidx))
    result = gs.get_values(grid_index=gidx)
    logger.info("")
    for i in result:
        logger.info ("result: {}".format(i))
    for i in gs.revision_list_mtime():
        logger.info ("revision: {}".format(i))

    expected = [[str(i*j) for i in range(1,10)] for j in range(1,10)]
    assert result == expected

    logger.info("")

    logger.info ("assert id_ori == gs.spreadsheet_cursor.id")
    assert id_ori == gs.spreadsheet_cursor.id

    ws_names_tmp = gs.worksheets(only='name')
    logger.info ("assert sorted(ws_names_tmp) == sorted(ws_names_ori)")
    assert sorted(ws_names_tmp) == sorted(ws_names_ori)
    logger.info ("assert ws_names_tmp == ws_names_ori")
    assert ws_names_tmp == ws_names_ori

    gs.remove_permission (email=demo_email)

    """
    cleanup
    """
    try:
        gsnew.delete_spreadsheet()
    except Exception as e:
        pass
        # logger.warning (e)

    try:
        gs.delete_spreadsheet()
    except Exception as e:
        logger.warning (e)
