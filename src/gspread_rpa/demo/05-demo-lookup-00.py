"""
High Level GoogleSheet API demo
"""

import os, sys
import logging
from logging.handlers import TimedRotatingFileHandler
from random import sample

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
        gs.delete_spreadsheet()
    except:
        pass

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
    gs.add_worksheet(title='lookup', rows=20 ,  cols=20,   tab_position=0)
    logger.info (gs.worksheet_cursor)

    data = [
        ['Patricio'  , 'Levi'     , 'Weiss'           ],
        ['Halimah'   , 'Pinocchio', 'Abraham'         ],
        ['Halimah'   , ''         , 'Abraham'         ],
        ['Diana'     , 'Yana'     , 'Hartley'         ],
        ['Margareta' , 'Nóirín'   , ''                ],
        ['Senan'     , 'Lihuén'   , ''                ],
        ['Tsvetana'  , 'Gislenus' , 'Salvatici'       ],
        ['Rava'      , 'Iosias'   , ''                ],
        ['Theotman'  , 'Baugulf'  , ''                ],
        ['Jákob'     , 'Derbiled' , 'Serra'           ],
        ['Günay'     , 'Rufina'   , 'Anker'           ],
        ['Terell'    , 'Raimonds' , 'Herbertsson'     ],
        ['Patrícia'  , 'Knut'     , 'Herrera'         ],
        ['Rohit'     , 'Suresh'   , 'Van Antwerpen'   ],
        ['Iunia'     , 'Jaromir'  , 'Hirano'          ],

        [''          , ''         , ''                ],

        ['Vanesa'    , 'Azra'     , 'Ubiña'           ],
        ['Sante'     , 'Phaidra'  , 'Tomić'           ],
        ['Eun - Ji'  , 'Kirabo'   , 'Zeman'           ],
        ['Susheela'  , 'Daniyyel' , 'Raptis'          ],
        ['Elisabet'  , 'Burkhart' , 'Kempf'           ],
        ['Ragnhildur', 'Jayanti'  , 'Christian'       ],
        ['Kumar'     , 'Millaray' , 'Van Alphen'      ],
        ['Valerianus', 'Matéo'    , 'Belluomo'        ],
        ['Ahmed'     , 'Raj'      , 'Vogt'            ],
        ['Irene'     , 'Diodorus' , 'Johansen'        ],
        [''          , 'Raj'      , 'Vogt'            ],
        ['Patrícia'  , ''         , 'Zeman'           ],
    ]

    def ltrim (l):
        if not l: return []
        s=-1
        for i,j in enumerate(l):
            if j and s<0: s = i
        return l[s:]

    def rtrim(l):
        if not l: return []
        e=-1
        for i,j in enumerate(reversed(l)):
            if j and e<0: e = i
        return l[:len(l) - e]

    def trim(l):
        return list(rtrim(ltrim(l)))

    gid = GridIndex(1,1,10,1)
    gs.update_cells(cells_index=gid, values=[[i for i in range(1, 10)]])
    gid = GridIndex(1,2,1,34)
    gs.update_cells(cells_index=gid, values=[[i for i in range(2, 35)]])

    logger.info ("data:\n {}".format("\n".join(["{}".format(i) for i in data])))

    csi = CellIndex(col=3, row=4)
    logger.info ("gs.update_cells(cells_index={}, values=data)".format(csi))
    gs.update_cells(cells_index=csi, values=data)

    result = gs.get_values()
    logger.info("")
    for i in result:
        logger.info ("result: {}".format(i))

    logger.info("")
    logger.info ("x search")
    for i in sample(data, len(data)):
        l = trim(i)
        logger.info ("")
        logger.info ("lookup: {}".format(l))
        match_location = gs.lookup_match (match=l, search_direction='x')
        logger.info (match_location)
        for m in match_location:
            r = gs.get_values (m)
            logger.info ("result: {}".format (r))
            try:
                assert l == trim(r[0] if r else ['']), "unexpected result {}".format(r)
            except AssertionError:
                logger.info ("{}, try next".format(r[0]))
                continue
            else:
                break
        else:
            raise ValueError

    logger.info ("")
    logger.info ("y search")
    ydata=list(zip(*data))
    for i in sample(ydata, len (ydata)):
        l = trim(i)
        logger.info ("")
        logger.info ("lookup: {}".format(l))
        match_location = gs.lookup_match (match=l, search_direction='y')
        logger.info (match_location)
        for m in match_location:
            r = gs.get_values (m)
            logger.info ("result: {}".format (trim([i[0] for i in r])))
            try:
                assert l == trim([i[0] for i in r]), "unexpected result {}".format(trim([i[0] for i in r]))
            except AssertionError:
                logger.info ("try next lookup {}".format(l))
                logger.info ("try next result {}".format(trim([i[0] for i in r]) if r else []))
                continue
            else:
                break
        else:
            raise ValueError
        logger.info ("")

    gs.remove_permission (email=demo_email)

    """
    cleanup
    """
    gs.delete_spreadsheet()
