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

    logger.info("")
    gidx = None
    logger.info("gs.get_values()".format(gidx))
    result = gs.get_values(grid_index=gidx)
    logger.info("")
    for i in result:
        logger.info ("result: {}".format(i))

    matches = ['[0-9][0-9]+']
    match_location = None

    logger.info ("")
    logger.info ("col lookup")
    logger.info ("")
    logger.info ("gs.lookup_match (match={}, search_direction='x'".format(matches))
    match_location = gs.lookup_match (match=matches, search_direction='x')
    logger.info ("Match location: {}".format(match_location if match_location else None))
    for i in match_location:
        vedic_9 = []
        data_at_loc = gs.get_values(i)
        logger.info ("get values at {}: {}".format(i, data_at_loc))
        for v in [i for j in data_at_loc for i in j]:
            tmp = 0
            for j in list(v):
                tmp += int(j)
            vedic_9.append(tmp)
        logger.info("gs.update_cells(cells_index={}, values=[{}])".format(i, vedic_9))
        gs.update_cells(cells_index=i, values=[vedic_9])
    result = gs.get_values()
    logger.info("")
    for i in result:
        logger.info ("result: {}".format(i))


    match_location = None
    logger.info ("")
    logger.info ("row lookup")
    logger.info ("")
    logger.info ("gs.lookup_match (match={}, search_direction='y'".format(matches))
    match_location = gs.lookup_match (match=matches, search_direction='y')
    logger.info ("Match location: {}".format(match_location if match_location else None))
    for i in match_location:
        vedic_9 = []
        data_at_loc = gs.get_values(i)
        logger.info ("get values at {}: {}".format(i, data_at_loc))
        for v in [i for j in data_at_loc for i in j]:
            tmp = 0
            for j in list(v):
                tmp += int(j)
            vedic_9.append(tmp)
        logger.info("gs.update_cells(cells_index={}, values=[{}])".format(i, vedic_9))
        gs.update_cells(cells_index=i, values=[vedic_9])

    result = gs.get_values()
    logger.info("")
    for i in result:
        logger.info ("result: {}".format(i))

    result = gs.get_values()
    expected = [
        ['1', '2', '3', '4', '5', '6', '7', '8', '9'],
        ['2', '4', '6', '8', '1', '3', '5', '7', '9'],
        ['3', '6', '9', '3', '6', '9', '3', '6', '9'],
        ['4', '8', '3', '7', '2', '6', '1', '5', '9'],
        ['5', '1', '6', '2', '7', '3', '8', '4', '9'],
        ['6', '3', '9', '6', '3', '9', '6', '3', '9'],
        ['7', '5', '3', '1', '8', '6', '4', '2', '9'],
        ['8', '7', '6', '5', '4', '3', '2', '1', '9'],
        ['9', '9', '9', '9', '9', '9', '9', '9', '9'],
    ]
    # for i in expected:
    #     logger.info ("expect: {}".format(i))

    assert result == expected

    gs.remove_permission (email=demo_email)

    """
    cleanup
    """
    logger.info("")
    gs.delete_spreadsheet()
