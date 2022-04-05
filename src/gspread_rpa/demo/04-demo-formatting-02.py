"""
High Level GoogleSheet API demo
"""

import os, sys
import logging
from logging.handlers import TimedRotatingFileHandler
from random import choices

if os.path.exists(
        os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'), 'gspread_rpa')):
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
from gspread_rpa import CellIndex, GridIndex, GoogleSheets, CellFormat, ColorMap

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

    gs.clear ([(1,1, 100, 100)])
    cf = CellFormat()
    cf.horizontal_alignment('left')
    cf.vertical_alignment('top')
    gs.prepare_cells_user_format (GridIndex(1,1,100,100), cf)
    gs.apply_cells_user_format()

    data = [[i*j for i in range(1,10)] for j in range(1,10)]
    logger.info ("data:\n {}".format("\n".join(["{}".format(i) for i in data])))

    # 1/4
    csi = CellIndex(col=1, row=1)
    logger.info ("gs.update_cells(cells_index={}, values=data)".format(csi))
    gs.update_cells(cells_index=csi, values=data)
    # 2/4
    csi = CellIndex(col=1, row=1)
    [i.reverse() for i in data]
    gs.update_cells(cells_index=GridIndex(start_col=10, start_row=1, end_col=18, end_row=9), values=data)
    # 3-4/4
    data = gs.get_values()
    data.reverse()
    gs.update_cells(cells_index=GridIndex(start_col=1, start_row=10, end_col=18, end_row=18), values=data)

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

    cf = CellFormat()
    cf.horizontal_alignment('center')
    cf.vertical_alignment('top')
    # cf.text.bold(False)
    # cf.text.font_size(8)
    # cf.text.foreground_color('black')
    gs.prepare_cells_user_format (GridIndex(1,1,20,20), cf)
    gs.apply_cells_user_format()

    for i in range(1, 10):
        matches = ["{}".format(i)]
        match_location = gs.lookup_match (match=matches, search_direction='x')
        for ml in match_location:
            logger.info ("match_location: {}".format(ml))

        color = ColorMap().names[i]
        cf = CellFormat()
        cf.horizontal_alignment('center')
        cf.vertical_alignment('top')
        cf.background_color(color)
        cf.text.foreground_color('white')
        cf.text.bold(True)
        cf.text.font_family("Arial")
        cf.text.font_size(12)
        cf.text.italic(False)
        if i == 9:
            cf.text_rotation.angle(90)
        for idx in match_location:
            gs.prepare_cells_user_format (idx, cf)
            # gs.apply_cells_user_format() # apply here to see the progress or less deeper for less request
        gs.apply_cells_user_format() # could be better applyed at the end to reduce the number of request
    # gs.apply_cells_user_format()

    for i in choices(list(range(1, 10)), k=10):
        c = ColorMap().names[i]
        match_location = gs.lookup_match (match=["{}".format(i)], search_direction='x')
        for k in choices(match_location, k=3):
            logger.info ("assert {} background_color_name == {}".format(k.start, c))
            gcf = gs.get_cell_user_format (k.start)
            assert gcf.background_color_name == c

    gs.remove_permission (email=demo_email)

    """
    cleanup
    """
    logger.info("")
    gs.delete_spreadsheet()
