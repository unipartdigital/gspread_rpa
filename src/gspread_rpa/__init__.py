import gspread
from gspread.utils import rowcol_to_a1, a1_to_rowcol, ValueRenderOption, ValueInputOption
from gspread.auth import local_server_flow
from gspread.auth import DEFAULT_SCOPES, DEFAULT_CREDENTIALS_FILENAME, DEFAULT_AUTHORIZED_USER_FILENAME
from .gspreadsheet_retry import SpreadsheetRetry, WorksheetRetry, ClientRetry
from .gspreadsheet_retry import exceptions, retry, error_quota_req
from .format_cell import CellFormat, ColorMap
import logging
from re import compile, IGNORECASE
from os import getenv, unlink, path

"""
GoogleSheets HighLevel wrapper around gspread
"""

__version__ = "1.0.2"

logger = logging.getLogger('GoogleSheets')


class CellIndex(object):
    def __init__(self, col=None, row=None):
        self.col=int(col) if col else None
        self.row=int(row) if row else None

    def __repr__(self):
        return "<{} col:{} row:{}>".format(
            self.__class__.__name__,
            self.col,
            self.row)

    def __hash__(self):
        return hash((self.col, self.row))

    def __eq__(self, other):
        return  self.col == other.col and self.row == other.row

    """ cell label in A1 notation, e.g. 'B1' """
    def from_a1(self, label):
        self.row, self.col =  a1_to_rowcol (label)
        return self

    def to_a1(self):
        result = rowcol_to_a1(col=self.col, row=self.row)
        return result

class GridIndex(object):
    def __init__(self, start_col=None, start_row=None, end_col=None, end_row=None):
        self.start = CellIndex(col=start_col, row=start_row)
        self.end = CellIndex(col=end_col, row=end_row)

    def __repr__(self):
        return "<{} start:{} end:{}>".format(
            self.__class__.__name__,
            repr(self.start), repr(self.end))

    def __hash__(self):
        return hash((self.start.col, self.start.row, self.end.col, self.end.row))

    def __eq__(self, other):
        return  self.start == other.start and self.end == other.end

class DataCache(object):

    def __init__(self):
        self.cached_cells = []
        self._expired = True

    def expired(self):
        return self._expired == True

    def close(self):
        self.cached_cells = []
        self._expired = True

    def store(self, data, cell_start=CellIndex(col=1, row=1)):
        self.cache_cells = [
            [
                gspread.Cell(
                    row=nrow + cell_start.row,
                    col=ncol + cell_start.col,
                    value=val
                ) for ncol,val in enumerate(row)
            ] for nrow,row in enumerate(data)]
        self._expired = False

class GoogleSheets(object):

    class AlreadyExists (Exception):
        pass

    class NotFound (Exception):
        pass

    class InitError (Exception):
        pass

    """
    run_mode: service for service account, user for oauth (require user interaction)
    if not set at object creation like gs = GoogleSheets(run_mode='service')
    will take the valie from GOOGLESHEETS_RUN_MODE env variable and default to 'service'
    """
    def __init__(self, run_mode='',
                 scopes=DEFAULT_SCOPES,
                 flow=local_server_flow,
                 credentials_filename=DEFAULT_CREDENTIALS_FILENAME,
                 authorized_user_filename=DEFAULT_AUTHORIZED_USER_FILENAME):
        assert run_mode in (None, '', 'service', 'user'), "run_mode set and not in 'service' or 'user'"
        self.run_mode = run_mode if run_mode else getenv('GOOGLESHEETS_RUN_MODE', 'service')
        self.spreadsheet_cursor = None
        self.worksheet_cursor = None
        self.cell_current_position = (1, 1)
        self.spreadsheet_revision = None
        if self.run_mode == 'service':
            self.gc = gspread.service_account()
            self.gc = ClientRetry(auth=self.gc.auth)
        else:
            for i in [1, 2]:
                self.gc = gspread.oauth(
                    scopes=scopes,
                    flow=flow,
                    credentials_filename=credentials_filename,
                    authorized_user_filename=authorized_user_filename
                )
                self.gc = ClientRetry(auth=gspread.auth.load_credentials())
                # validate grant by attempting a request
                try:
                    tmp = self.gc.list_spreadsheet_files(title='')
                except Exception as e:
                    logger.error (dir(e))
                    if 'RefreshError' in type(e).__name__:
                        if path.exists (authorized_user_filename):
                            unlink (authorized_user_filename)
                            continue
                    else:
                        raise e
                else:
                    break
            else:
                raise self.InitError from None
        self.client_ext = ClientRetry(self.gc.auth, self.gc.session)
        self.placeholder = []
        self.data_cache = DataCache()

    """
    id
    """
    def spreadsheet_id(self):
        if self.spreadsheet_cursor:
            return self.spreadsheet_cursor.id

    """
    title
    """
    def spreadsheet_title(self):
        if self.spreadsheet_cursor:
            return self.spreadsheet_cursor.title

    """
    Creates a new spreadsheet.
    """
    def create(self, title, folder_id=None):
        if self.spreadsheet_cursor is None:
            self.spreadsheet_cursor = SpreadsheetRetry(self.gc.create(title=title, folder_id=folder_id))
            logger.info ("create: {}".format(self.spreadsheet_cursor))


    """
    Deletes a spreadsheet.
    """
    def delete_spreadsheet(self):
        logger.info ("delete: {}".format(self.spreadsheet_cursor))
        if self.spreadsheet_cursor is not None:
            try:
                self.gc.del_spreadsheet(self.spreadsheet_cursor.id)
            except Exception as e:
                logger.error ("delete_spreadsheet: {}".format(e))
                raise e
            else:
                self.close()
        else:
            raise self.NotFound ("delete spreadsheet not open/created")

    """
    Give permission to a spreadsheet

    Note: remeber to give some self permission to your own email if using a service account.
    Example:
     Give Otto a write permission on this spreadsheet:
      gs.give_permission('otto@example.com', perm_type='user', role='writer')
     Transfer ownership to Otto:
      gs.give_permission('otto@example.com', perm_type='user', role='owner')
    """
    def give_permission (self, email, perm_type, role, notify=False, email_message=None, with_link=False):
        assert perm_type in (
            'user', 'group', 'domain', 'anyone'
        ), "Allowed perm_type are: user, group, domain, anyone."
        assert role in ('owner', 'writer', 'reader'), "Allowed role are: owner, writer, reader"
        if self.spreadsheet_cursor is not None:
            self.gc.insert_permission(file_id=self.spreadsheet_cursor.id,
                                      value=email, perm_type=perm_type, role=role, notify=notify,
                                      email_message=email_message, with_link=with_link)

    """
    Remove permission

    Remove Otto's write permission for this spreadsheet
     gs.remove_permission ('otto@example.com', role='writer')
    Remove all Otto's permissions for this spreadsheet
     gs.remove_permission ('otto@example.com')
    """
    def remove_permission(self, email, role='any'):
        assert role in ('any', 'owner', 'writer', 'reader'), "Allowed role are: owner, writer, reader"
        if self.spreadsheet_cursor is not None:
            self.spreadsheet_cursor.remove_permissions(value=email, role=role)

    """
    List Permission

    example:
       for p in gs.list_permission():
        print ("permission: {}".format([p[i] for i in ['type', 'role', 'emailAddress']]))
    """
    def list_permission(self):
        if self.spreadsheet_cursor is not None:
            return self.spreadsheet_cursor.list_permissions()

    """
    """
    def is_open(self):
        return self.spreadsheet_cursor is not None

    """
    may be used to open an other spreadsheet reusing an existing GoogleSheet instance
    """
    def close(self):
        self.spreadsheet_cursor = None
        self.worksheet_cursor = None
        self.cell_current_position = (1, 1)
        self.spreadsheet_revision = None
        self.data_cache.close()

    """
    Open a spreadsheet try in order 'url', id and then title
    if nay tab_name, tab_position, tab_id is set open the worsheet trying first
    'tab_id', 'tab_name', 'tab_position' and then tab '0'
    """
    @retry(tries=15, delay=5, backoff=2, except_retry=[error_quota_req])
    def open (self, title=None, url=None, key=None, tab_name=None, tab_position=None, tab_id=None):
        if self.spreadsheet_cursor is None:
            assert any ([title, url, key]), "opening a spreadsheet require a title, an url or a key"
            for i in ['url', 'key', 'title']:
                try:
                    if i == 'url' and url:
                        self.spreadsheet_cursor = SpreadsheetRetry(self.gc.open_by_url(url))
                    if i == 'key' and key:
                        self.spreadsheet_cursor = SpreadsheetRetry(self.gc.open_by_key(key))
                    if i == 'title' and title:
                        self.spreadsheet_cursor = SpreadsheetRetry(self.gc.open(title))
                    if self.spreadsheet_cursor is None:
                        continue
                    else:
                        break
                except exceptions.SpreadsheetNotFound as e:
                    continue
                except Exception as e:
                    logger.error (e)
                    raise
                else:
                    break

            if not self.spreadsheet_cursor: raise self.NotFound
            logger.info ("open spreadsheet: {}".format(self.spreadsheet_cursor))
            try:
                self.spreadsheet_revision = self.client_ext.revision_last(
                    self.spreadsheet_cursor.id, revision_id=None)
            except exceptions.APIError as e:
                if hasattr(e, 'code') and e.code == 403:
                    logger.error (e)
            except Exception as e:
                logger.error (e)
                raise e
            else:
                pass
                # logger.info ("spreadsheet_revision: {}".format(self.spreadsheet_revision))
            self.data_cache.close()
            if not any ([tab_name, tab_position, tab_id]): return

        self.worksheet_cursor = None
        for i in ['tab_id', 'tab_name', 'tab_position']:
            try:
                if i == 'tab_name' and tab_name:
                    logger.debug ("open tab name {}".format(tab_name))
                    self.worksheet_cursor = [
                        w for w in [
                            w for w in self.worksheets()
                        ] if w.title.strip().lower() in [
                            n.strip().lower() for n in [tab_name]]]
                    assert self.worksheet_cursor, "error in open tab name: {}".format(tab_name)
                    self.worksheet_cursor = self.worksheet_cursor[0]
                if i == 'tab_id' and tab_id:
                    tab_id = int (tab_id)
                    logger.debug ("open tab id {}".format(tab_id))
                    self.worksheet_cursor = [
                        w for w in [
                            w for w in self.worksheets()
                        ] if w.id in [tab_id]]
                    assert self.worksheet_cursor, "error in open tab id: {}".format([tab_id])
                    self.worksheet_cursor = self.worksheet_cursor[0]
                if i == 'tab_position' and tab_position:
                    logger.debug ("open tab pos {}".format(tab_position))
                    self.worksheet_cursor = [
                        (i,w) for (i,w) in
                        enumerate(self.worksheets()) if i == int(tab_position)]
                    assert self.worksheet_cursor, "error in open tab position: {}".format(tab_position)
                if self.worksheet_cursor is None:
                    continue
                else:
                    break
            except AssertionError as e:
                logger.error (e)
                continue
            except Exception as e:
                logger.error (e)
            else:
                break
        else:
            logger.debug ("warning open tab pos {}".format(0))
            self.worksheet_cursor = [
                (i,w) for (i,w) in
                enumerate(self.worksheets()) if i == int(0)]
            assert self.worksheet_cursor, "error in open tab position: {}".format(0)
        self.data_cache.close()

    """
    return a list with all the fields for each revision available
    """
    def revision_list(self):
        result = self.client_ext.revision_list(spreadsheet_id=self.spreadsheet_cursor.id)
        return result

    """
    return a list of namedtuple 'IdModifiedTime' i.id, i.mtime
    of all the revision currently available sorted by mtime
    """
    def revision_list_mtime (self):
        result = self.client_ext.revision_list_mtime(spreadsheet_id=self.spreadsheet_cursor.id)
        return result

    """
    given one of the supported format extension,
    return a mime type suitable for the file _export function
    """
    def ext2mime (self, extension):
        mimes = {
            'csv': 'text/csv',
            'ods': 'application/vnd.oasis.opendocument.spreadsheet',
            'pdf': 'application/pdf',
            'tsv': 'text/tab-separated-values',
            'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'zip': 'application/zip',
        }
        logger.debug ("ext2mime: {}".format(extension.lower()))
        if extension.lower() in mimes:
            return mimes[extension.lower()]
        else:
            raise self.NotFound ("no mime map for {}".format(extension))

    """
    Export a specific revision, or last modified if revision_id=None
    usage:
       with open ('/tmp/demo-revision-1.pdf', 'wb') as fd:
        gs.file_export (fd, revision_id=1, mime_type='application/pdf')
    """
    def file_export(self, fd, revision_id='head', mime_type=None, extension=None):
        if not mime_type and extension:
            mime_type = self.ext2mime(extension)
        elif not mime_type and not extension:
            if fd.name:
                mime_type = self.ext2mime(fd.name.rpartition('.')[-1])
        assert mime_type, "unknow mime_type"
        result = self.client_ext.file_export(
            fd, spreadsheet_id=self.spreadsheet_cursor.id, revision_id=revision_id, mime_type=mime_type)
        return result

    """
    upload content of file descripor fd on success return new GoogleSheets if return_object == True
    else a file ressource id like '1qhNTwrt6BGcOX3c3DqaINMLeIsc-ceHJ'
    """
    def file_upload(self, fd, title='', mime_type=None, extension=None,
                    return_object=True):
        if extension:
            mime_type= self.ext2mime(extension)
        new_id = self.client_ext.file_upload(fd, title=title, mime_type=mime_type)
        if return_object:
            new_object = GoogleSheets(run_mode=self.run_mode)
            new_object.open(key=new_id)
            return new_object
        return new_id

    """
    backup into tmp, then remove all the worksheet from self and then
    copy all the worksheet from src to self
    self should keep the same ID but worksheets may have differents ID.
    worksheets names should be consistant.

    delete remove src spreadsheet at the end
    """
    def overwrite(self, src, delete=True):
        assert isinstance(src, GoogleSheets), "source not a GoogleSheets instance, {}".format(src.__class__)
        assert src.spreadsheet_cursor, "src not open {}".format(src.spreadsheet_cursor)
        try:
            tmp =  GoogleSheets(run_mode=self.run_mode)
            tmp.create (title="{}.bak".format(self.spreadsheet_cursor.title))
            wid_ori = self.worksheets(only='id')
            """ backup """
            for w in self.worksheets():
                w.copy_to(tmp.spreadsheet_cursor.id)
            """ spreadsheet must have at least one worksheet """
            self.add_worksheet(title="{}".format(self))
            """ ori clean up"""
            for w in self.worksheets():
                if w.id in wid_ori:
                    try:
                        self.spreadsheet_cursor.del_worksheet(w)
                    except Exception as e:
                        err_name = "{}.{}".format(e.__class__.__module__ ,  e.__class__.__name__)
                        err_code = None
                        err_code = [i['code'] for i in e.args if 'code' in i]
                        err_code = int(err_code[0]) if err_code else None
                        logger.debug ("exception handling {} {}".format(err_name, err_code))
                        if err_code == 400:
                            # handle protected cell or object
                            logger.warning ("L434: {}".format(e))
                        else:
                            raise e
            """ cp from src to self """
            for w in src.worksheets():
                logger.info ("COPY {}.copy_to({})".format(w, self.spreadsheet_cursor.id))
                w.copy_to(self.spreadsheet_cursor.id)
            """ copy prepand with 'Copy of' """
            for w in self.worksheets():
                try:
                    logger.info (
                        "UPDATE {}.update_title({})".format(w, w.title.replace("Copy of", "").strip())
                    )
                    w.update_title(w.title.replace("Copy of", "").strip())
                except Exception as e:
                    err_name = "{}.{}".format(e.__class__.__module__ ,  e.__class__.__name__)
                    err_code = None
                    err_code = [i['code'] for i in e.args if 'code' in i]
                    err_code = int(err_code[0]) if err_code else None
                    logger.debug ("exception handling {} {}".format(err_name, err_code))
                    if err_code == 400:
                        # handle protected cell or object
                        logger.warning ("L456: {}".format(e))
                    else:
                        raise e
            self.open(tab_name="{}".format(self))
            self.delete_worksheet()
        except Exception as e:
            logger.error ("overwrite: {}".format(e))
            logger.error ("copy available: {}".format(tmp))
            raise e
        else:
            tmp.delete_spreadsheet()
            if delete: src.delete_spreadsheet()

    """
    delete previously uploaded user file. return True on success
    """
    def file_delete(self, id):
        result = self.client_ext.file_delete(id=id)
        return result

    """
    Add a new worksheet (Tab) into the spreadsheet
     the newly created workseet become the active worksheet, use open to switch to an other one
    """
    def add_worksheet(self, title, cols=26, rows=56, tab_position=None, raise_if_exists=False):
        try:
            self.worksheet_cursor = WorksheetRetry(self.spreadsheet_cursor.add_worksheet(
                title=title, rows=rows, cols=cols, index=tab_position))
        except exceptions.APIError as e:
            if [i['code'] == 400 for i in e.args if 'code' in i] and [
                    'A sheet with the name' in i['message'] or
                    'already exists' in i['message'] for i in e.args if 'message' in i]:
                if raise_if_exists:
                    raise self.AlreadyExists ("{}".format(title)) from None
                else:
                    logger.debug ("{} {}".format(e, repr(e)))
                    self.worksheet_cursor = [
                        w for w in [
                            w for w in self.worksheets()
                        ] if w.title.strip().lower() == title.strip().lower()]
                    assert self.worksheet_cursor, "error in open tab name: {}".format(title)
                    self.worksheet_cursor = self.worksheet_cursor[0]
            else: raise e
        else:
            self.data_cache.close()
            logger.info("create {}".format(self.worksheet_cursor))

    """
    return a list of titles (only='title'), id (only='id') or of the whole object
    """
    @retry(tries=15, delay=5, backoff=2, except_retry=[error_quota_req])
    def worksheets(self, only=None):
        if only is None:
            return [WorksheetRetry(w) for w in self.spreadsheet_cursor.worksheets()]
        else:
            return [getattr(i, only) for i in self.spreadsheet_cursor.worksheets() if hasattr (i, only)]

    """
    """
    def reorder_worksheets(self, worksheets_in_desired_order):
        logger.info("reorder_worksheets: {}".format(worksheets_in_desired_order))
        self.spreadsheet_cursor.reorder_worksheets(worksheets_in_desired_order)

    """
    Delete the active worksheet
    """
    def delete_worksheet(self):
        assert self.worksheet_cursor is not None, "no active worksheet to delete"
        logger.info ("delete {}".format(self.worksheet_cursor))
        self.worksheet_cursor = self.spreadsheet_cursor.del_worksheet(self.worksheet_cursor)
        self.data_cache.close()
        self.worksheet_cursor = None

    """
    resize the worksheet to cols, rows
    """
    def resize(self, cols=None, rows=None):
        assert self.worksheet_cursor is not None, "no active worksheet to resize"
        self.worksheet_cursor.resize(cols=cols, rows=rows)
        self.data_cache.close()
        logger.info ("{} col_count={} row_count={}".format(
            self.worksheet_cursor, self.worksheet_cursor.col_count, self.worksheet_cursor.row_count))

    """
    clear all cached data
    """
    def close_cache (self):
        self.data_cache.close()

    """
    lookup match in search_direction  X (col) or Y (row)
    return a list of GridIndex (start.col, start.row, end.col, end.row) if match else None
    the result list is sorted with the longest at the end

    as the funtion use regexpr it may be needed to validate by fetching the data
    """
    def lookup_match (self, match=[], search_direction='col', default_regex=r"\b({})\b"):
        assert self.worksheet_cursor, "worksheet not open"

        if self.data_cache.expired():
            data = self.get_values()
            self.data_cache.store (data)

        rs = ""
        for i in match[:-1]:
            rs += default_regex.format(i) + "|" if i else ''
        rs += default_regex.format(match[-1]) if match and match[-1] else '(^$)'
        logger.info ("r: {}".format(rs))
        rc = compile(rs, IGNORECASE)

        cell_find_list = []
        for i in self.data_cache.cache_cells:
            for j in i:
                if rc.search (j.value):
                    cell_find_list.append (j)
                elif j.value == '':
                    if cell_find_list and search_direction in ('col', 'x') and (
                            cell_find_list[-1].row == j.row and cell_find_list[-1].col == j.col - 1):
                        cell_find_list.append (j)
                    elif cell_find_list and search_direction in ('row', 'y') and (
                            cell_find_list[-1].col == j.col and cell_find_list[-1].row == j.row - 1):
                        cell_find_list.append (j)

        logger.debug("find: {}".format(cell_find_list))
        if search_direction.lower() in ('col', 'x'):
            cell_find_list.sort(key=lambda x: (int(x.row), int(x.col)), reverse=False)
        else:
            cell_find_list.sort(key=lambda x: (int(x.col), int(x.row)), reverse=False)

        result=[]
        ridx=GridIndex()
        for cur, nxt in zip(cell_find_list, cell_find_list[1:] + [CellIndex()]):
            # logger.info ("cur: {} nxt: {}".format(cur, nxt))
            ridx.start = cur if (ridx.start.row, ridx.start.col) == (None, None) else ridx.start
            if  cur.col == nxt.col and cur.row + 1 == nxt.row:
                continue
            elif  cur.row == nxt.row and cur.col + 1 == nxt.col:
                continue
            else:
                ridx.end = cur
                result.append (GridIndex(ridx.start.col, ridx.start.row, ridx.end.col, ridx.end.row))
                ridx = GridIndex()
        if search_direction.lower() in ('col', 'x'):
            result.sort(key=lambda x: (x.end.col - x.start.col), reverse=False)
        else:
            result.sort(key=lambda x: (x.end.row - x.start.row), reverse=False)
        logger.debug ("lookup_match result: {}".format(result))
        return result

    """
    get values from column col (start index 1)
    """
    def get_values_col (self, col, **kwargs):
        assert self.worksheet_cursor, "worksheet not open"
        result = self.worksheet_cursor.col_values(col, **kwargs)
        return result

    """
    get values from line row (start index 1)
    """
    def get_values_row (self, row, **kwargs):
        assert self.worksheet_cursor, "worksheet not open"
        result = self.worksheet_cursor.row_values(row, **kwargs)
        return result

    """
    Returns a list of lists containing all values from specified range
    get values from range GridIndex or tuple (start_col, start_row, end_col, end_row)
    if grid_index is not defined, returns values from all non empty cells
    """
    def get_values (self, grid_index=None, **kwargs):
        assert self.worksheet_cursor, "worksheet not open"
        if isinstance(grid_index, GridIndex):
            s = rowcol_to_a1(col=grid_index.start.col, row=grid_index.start.row)
            e = rowcol_to_a1(col=grid_index.end.col,   row=grid_index.end.row)
            range_name = "{}:{}".format(s, e)
        elif isinstance(grid_index, tuple) and len(grid_index) == 4:
            s = rowcol_to_a1(col=grid_index[0], row=grid_index[1])
            e = rowcol_to_a1(col=grid_index[2], row=grid_index[3])
            range_name = "{}:{}".format(s, e)
        else:
            range_name = None
        result = self.worksheet_cursor.get_values(range_name=range_name, **kwargs)
        return result

    """
    delete cols from start to end
    """
    def delete_cols(self, start_index, end_index=None):
        assert self.worksheet_cursor, "worksheet not open"
        self.worksheet_cursor.delete_columns(start_index, end_index=end_index)
        self.data_cache.close()

    """
    delete rows from start to end
    """
    def delete_rows(self, start_index, end_index=None):
        assert self.worksheet_cursor, "worksheet not open"
        self.worksheet_cursor.delete_rows(start_index, end_index=end_index)
        self.data_cache.close()

    """
    update a single cell value
    """
    def update_cell(self, col, row, value):
        assert self.worksheet_cursor, "worksheet not open"
        self.worksheet_cursor.update_cell(col=col, row=row, value=value)
        self.data_cache.close()

    """
    Clears multiple ranges in one API call
    example clear from col 1 row 1 to col 3 row 3 and
               as well col 7 row 1 to col 8 row 1 and a single cell (col 4, row 5) too
            clear([(1,1,3,3), (7,1,8,1), (4, 5)]
    """
    def clear (self, grid_index=[], **kwargs):
        assert self.worksheet_cursor, "worksheet not open"
        range_list = []
        for grid_idx in grid_index:
            if isinstance(grid_idx, GridIndex):
                s = rowcol_to_a1(col=grid_idx.start.col, row=grid_idx.start.row)
                e = rowcol_to_a1(col=grid_idx.end.col,   row=grid_idx.end.row)
                range_list.append("{}:{}".format(s, e))
            elif isinstance(grid_idx, tuple) and len(grid_idx) == 4:
                s = rowcol_to_a1(col=grid_idx[0], row=grid_idx[1])
                e = rowcol_to_a1(col=grid_idx[2], row=grid_idx[3])
                range_list.append("{}:{}".format(s, e))
            elif isinstance(grid_idx, tuple) and len(grid_idx) == 2:
                s = rowcol_to_a1(col=grid_idx[0], row=grid_idx[1])
                range_list.append("{}:{}".format(s, s))
            elif isinstance(grid_idx, CellIndex):
                s = rowcol_to_a1(col=grid_idx.col, row=grid_idx.row)
                range_list.append("{}:{}".format(s, s))
            else:
                raise ValueError ("clear idx {}".format(grid_idx))
        result = self.worksheet_cursor.batch_clear(ranges=range_list, **kwargs)
        self.data_cache.close()
        return result

    """
    update from the matrix (list of list) values starting at the cell_index location
                                              A B C D
    [ [1,2,3], [4,5,6] ] at (col 2, row 1) ->   1 2 3
                                                4 5 6

    gs.update_cells(CellIndex(col=2, row=1), values=[ [1,2,3], [4,5,6] ])
    to work on row direction provide correct array or use GridIndex or else use transpose = True
    ex:
    gs.update_cells(cells_index=<GridIndex start:<CellIndex col:7 row:4> end:<CellIndex col:7 row:8>>,
                    values=[[1, 8, 6, 4, 2]])
    use of Full GridIndex ok
    gs.update_cells(cells_index=(col:7, row=4),
                    values=[[1, 8, 6, 4, 2]])
    KO use transpose=True or pass the data as : [[1], [8], [6], [4], [2]]
    """
    def update_cells(self, cells_index, values, transpose=False,
                     value_input_option=ValueInputOption.user_entered,**kwargs):
        assert self.worksheet_cursor, "worksheet not open"
        start_col = 0
        start_row = 0
        if isinstance(cells_index, GridIndex):
            start_col = cells_index.start.col
            start_row = cells_index.start.row
        elif isinstance(cells_index, CellIndex):
            start_col = cells_index.col
            start_row = cells_index.row
        elif isinstance(cells_index, tuple) and len(cells_index) == 2:
            start_col = cells_index[0]
            start_row = cells_index[1]
        else:
            raise ValueError ("cells_index {}".format(cells_index))
        s = rowcol_to_a1(col=start_col, row=start_row)
        if  isinstance(cells_index, GridIndex):
            e = rowcol_to_a1(col=cells_index.end.col, row=cells_index.end.row)
        else:
            values = [list(sublist) for sublist in list(zip(*values))] if transpose else values
            e = rowcol_to_a1(col=max(map(len, values)) + start_col - 1, row=len (values) + start_row - 1)
        range_name = "{}:{}".format(s, e)
        cell_list = self.worksheet_cursor.range(range_name)
        idx=0
        for c in values:
            for v in c:
                cell_list[idx].value = v
                idx += 1
        result = self.worksheet_cursor.update_cells(
            cell_list=cell_list, value_input_option=value_input_option, **kwargs)
        self.data_cache.close()
        return result

    """
    refresh_ref
    try to refresh the reference in a spreadsheet by overwriting the same formula
    return the number of reference still unresoved in the spreadsheet
    """
    def refresh_ref (self):
        ref_count = 0
        t = self.spreadsheet_title()
        for w in self.worksheets():
            logger.info ("s {} w {}".format(t, w))
            self.open(tab_id=w.id)
            match_location = self.lookup_match (match=['#REF!'], default_regex=r"^{}$")
            for m in match_location:
                logger.info ("refresh {}: {}".format(w, m))
                v = self.get_values(m, value_render_option=ValueRenderOption.formula)
                try:
                    self.update_cells (m, v, value_input_option=ValueInputOption.user_entered)
                except Exception as e:
                    logger.warning ("refresh_ref {}".format(e))
            match_location = self.lookup_match (match=['#REF!'], default_regex=r"^{}$")
            for m in match_location:
                ref_count += 1
                logger.warning ("refresh ref unresolved {} {}".format(w, m))
        if ref_count > 0: logger.info ("refresh_ref {} unresolved ref in {}".format(ref_count, t))
        return ref_count


    """
    Cell Formatting
    """

    """
    return a CellFormat object, the format the user entered for the cell at cell_index.
    """
    def get_cell_user_format (self, cell_index):
        assert isinstance(cell_index, CellIndex)
        s = rowcol_to_a1(col=cell_index.col, row=cell_index.row)
        range = "'{}'!{}".format (self.worksheet_cursor.title, s)
        logger.debug ("get_cell_format: {}".format(range))
        resp = self.spreadsheet_cursor.fetch_sheet_metadata({
            'includeGridData': True,
            'ranges': [range],
            'fields': 'sheets.data.rowData.values.userEnteredFormat'})
        logger.debug ("get_cell_format: {}".format(resp))
        data = resp['sheets'][0]['data'][0]
        if 'rowData' in data:
            return CellFormat().dict2o(data['rowData'][0]['values'][0]['userEnteredFormat'])
        return CellFormat().dict2o({})

    """
    return a list of list of CellFormat object if
    the format the user entered for the cells at grid_index range exist otherwise ''
    """
    def get_cells_user_format (self, grid_index):
        assert isinstance(grid_index, GridIndex)
        s = rowcol_to_a1(col=grid_index.start.col, row=grid_index.start.row)
        e = rowcol_to_a1(col=grid_index.end.col, row=grid_index.end.row)
        range = "'{}'!{}:{}".format (self.worksheet_cursor.title, s, e)
        logger.info ("get_cell_format: {}".format(range))
        result =  []
        resp = self.spreadsheet_cursor.fetch_sheet_metadata({
            'includeGridData': True,
            'ranges': [range],
            'fields': 'sheets.data.rowData.values.userEnteredFormat'})
        for data in resp['sheets']:
            for row_data in data['data']:
                if 'rowData' in row_data:
                    for values in row_data['rowData']:
                        result_col = []
                        if 'values' in values:
                            for userEnteredFormat in values['values']:
                                if 'userEnteredFormat' in userEnteredFormat:
                                    tmp = CellFormat().dict2o(userEnteredFormat['userEnteredFormat'])
                                    result_col.append (tmp)
                                else:
                                    result_col.append ('')
                        result.append (result_col)
            # for r in result:
            #     logger.info (r)
            #     for c in r:
            #         if isinstance(c, CellFormat):
            #             logger.info(c.o2dict())
        return result

    def prepare_cells_user_format (self, grid_index, cell_format):
        self.placeholder.append((grid_index, cell_format))

    def cancel_cells_user_format (self, grid_index, cell_format):
        self.placeholder = []

    def apply_cells_user_format (self):
        requests = []
        body = {}
        for i in self.placeholder:
            idx = i[0]
            fmt = i[1]
            repeat_cell = {}
            range = {}
            range['range'] = {}
            range['range']['sheetId'] = self.worksheet_cursor.id
            if idx.start.row and idx.start.row > 0:
                range['range']['startRowIndex'] = idx.start.row - 1
            if idx.end.row and idx.end.row >= 1:
                range['range']['endRowIndex'] = idx.end.row
            if idx.start.col and idx.start.col > 0:
                range['range']['startColumnIndex'] = idx.start.col - 1
            if idx.end.col and idx.end.col >= 1:
                range['range']['endColumnIndex'] = idx.end.col
            cell = {}
            cell['cell'] = {}
            cell['cell']['userEnteredFormat'] = fmt.o2dict()
            cell.update ({'fields' : 'userEnteredFormat'})
            repeat_cell['repeatCell'] = dict (range)
            repeat_cell['repeatCell'].update (cell)
            requests.append (repeat_cell)
        if requests == []: return {}
        body['requests'] = requests
        body.update ({'includeSpreadsheetInResponse': False})
        body.update ({'responseRanges': []})
        body.update ({'responseIncludeGridData': False})
        logger.debug ("apply_cells_user_format: {}".format(body))
        result = self.spreadsheet_cursor.batch_update(body)
        logger.info (result)
        self.placeholder = []
        return result
