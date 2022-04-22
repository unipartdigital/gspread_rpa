from gspread import Spreadsheet, Worksheet, utils, exceptions, Client
from gspread.urls import DRIVE_FILES_API_V3_URL
import logging
from types import SimpleNamespace
from datetime import datetime
import os, tempfile
from collections import namedtuple
# from gspread_formatting import functions
import json
from .retry import retry

logger = logging.getLogger('gspreadsheet_retry')

"""
 extend some of  gspread classes
"""

RetryException = namedtuple ('RetryException', ['name', 'code'])

"""
require a fully qualified exception class name
"""
error_quota_req = RetryException(name='gspread.exceptions.APIError', code=429)
error_quota_qps = RetryException(name='gspread.exceptions.APIError', code=403)

class ClientRetry(Client):

    def __init__(self, auth, session=None):
        super(type(self), self).__init__(auth=auth, session=session)

    # @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_qps])
    # def open(self, title, folder_id=None):
    #     return super(type(self), self).open(title=title, folder_id=folder_id)

    # @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_qps])
    # def open_by_key(self, key):
    #     return super(type(self), self).open_by_key(key=key)

    # @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_qps])
    # def open_by_url(self, url):
    #     return super(type(self), self).open_by_url(url=url)

    # @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_qps])
    # def openall(self, title=None):
    #     return super(type(self), self).openall(title=title)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_qps, error_quota_req])
    def create(self, title, folder_id=None):
        logger.info ("Client create: {}".format(self))
        return super(type(self), self).create(title=title, folder_id=folder_id)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_qps])
    def revision_list (self, spreadsheet_id):
        # [
        #     {'kind': 'drive#revision', 'id': '1', 'mimeType': 'application/vnd.google-apps.spreadsheet',
        #      'modifiedTime': '2021-11-15T09:25:19.072Z'
        #      },
        #     {'kind': 'drive#revision', 'id': '16', 'mimeType': 'application/vnd.google-apps.spreadsheet',
        #      'modifiedTime': '2021-12-10T04:23:59.719Z'
        #      },
        #     {'kind': 'drive#revision', 'id': '40', 'mimeType': 'application/vnd.google-apps.spreadsheet',
        #      'modifiedTime': '2021-12-13T11:42:10.496Z'}
        # ]
        revisions = []
        page_token = ""
        url = "{}/{}/revisions".format(DRIVE_FILES_API_V3_URL, spreadsheet_id)
        fields = '*'
        params = {'fields': fields}

        while page_token is not None:
            if page_token:
                params["pageToken"] = page_token
            res = self.request("get", url, params=params).json()
            revisions.extend(res["revisions"])
            page_token = res.get("nextPageToken", None)
        return revisions

    """
    return a list of namedtuple 'IdModifiedTime' i.id, i.mtime
    of all the revision currently available sorted by mtime
    """
    def revision_list_mtime (self, spreadsheet_id):
        IdMtime = namedtuple('IdModifiedTime', ['id', 'mtime'])
        rl = [IdMtime(i['id'], i['modifiedTime']) for i in self.revision_list(spreadsheet_id)]
        rl = rl if rl else []
        return rl

    """
    return last revision or sprecified revision_id if found
    """
    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_qps])
    def revision_last(self, spreadsheet_id, revision_id=None):
        rev = self.revision_list (spreadsheet_id)
        rev = [SimpleNamespace(**n) for n in rev]
        rev_by_time = sorted(rev, key=lambda x:
                             datetime.strptime(getattr(x, 'modifiedTime'), "%Y-%m-%dT%H:%M:%S.%fZ"),
                             reverse=True)
        if revision_id is None or revision_id == 'head':
            last = rev_by_time[0] if len (rev_by_time) >= 1 and hasattr(rev_by_time[0], 'id') else None
            return last

        last = [n for n in rev_by_time if hasattr(n, 'id') and int(n.id) == int(revision_id)]
        last = last[0] if len (last) > 0 else None
        return last

    """
    usage:
       with open ('/tmp/demo.ods', 'wb') as fd:
        gs.file_export (fd, gs.spreadsheet_cursor.id, revision_id=1,
                   mime_type='application/x-vnd.oasis.opendocument.spreadsheet')

    """
    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_qps])
    def file_export(self, fd, spreadsheet_id, revision_id='head',
                    mime_type='application/x-vnd.oasis.opendocument.spreadsheet'):
        revision = self.revision_last(spreadsheet_id, revision_id)
        assert revision is not None
        params= {}
        export_link=revision.exportLinks[mime_type]
        try:
            os.fstat(fd.fileno())
            fd.seek(0)
        except Exception as e:
            logger.error (e)
            raise
        else:
            res = self.request("get", export_link, params=params)
            for chunk in res.iter_content(chunk_size=128):
                fd.write(chunk)
            logger.info ("file_export: fname={} fsize={}".format(fd.name, fd.tell()))

    """
    upload and convert content of file descripor fd on success return a file id
    {
      'kind': 'drive#file',
      'id': '1LPrKh6HgZ0LOmc3C3ZrOT2Tef_HIhsipjDAkHdOaQrI',
      'name': 'Untitled',
      'mimeType': 'application/vnd.google-apps.spreadsheet'
    }
    """
    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_qps])
    def file_upload(self, fd, title='', mime_type='application/x-vnd.oasis.opendocument.spreadsheet'):
        headers = None
        params = {
            "uploadType": "multipart",
        }
        data = {
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        }
        if title:
            data['name'] = title
        DRIVE_FILES_API_V3_UPLOAD_URL = "https://www.googleapis.com/upload/drive/v3/files"
        try:
            fd.seek(0)
        except Exception as e:
            logger.error (e)
            raise
        else:
            res = self.request(method="post",
                               endpoint="{}".format(DRIVE_FILES_API_V3_UPLOAD_URL),
                               params=params,
                               files={
                                   'meta': ('body', json.dumps(data), 'application/json'),
                                   'file': (os.path.basename (fd.name),fd, mime_type)
                               },
                               headers=headers)
            assert res.status_code == 200
            result = res.json()
            logger.info ("file_upload: {}".format(result))
            new_id = result['id'] if 'id' in result else None
            return new_id

    """
    delete previously uploaded user file. return True on success
    """
    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_qps])
    def file_delete (self, id):
        params = {}
        j = {}
        logger.info ("delete {}/{}".format(DRIVE_FILES_API_V3_URL, id))
        res = self.request("delete", "{}/{}".format(DRIVE_FILES_API_V3_URL, id))
        return res.status_code in (200, 204)

class SpreadsheetRetry(Spreadsheet):

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def __init__(foo, self):
        super().__init__(self.client, properties=self._properties)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def worksheets(self):
        return super(type(self), self).worksheets()

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def batch_update(self, body):
        return super(type(self), self).batch_update(body=body)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def fetch_sheet_metadata(self, params=None):
        return super(type(self), self).fetch_sheet_metadata(params=params)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def _spreadsheets_sheets_copy_to(self, sheet_id, destination_spreadsheet_id):
        return super(type(self), self)._spreadsheets_sheets_copy_to(sheet_id, destination_spreadsheet_id)

    # @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    # def del_worksheet(self, worksheet):
    #     logger.error ("del {}".format(self))
    #     return super(type(self), self).del_worksheet(self)

    # """
    # gspread_formatting
    # """
    # @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    # def get_default_format(self):
    #     return functions.get_default_format(spreadsheet=self)


class WorksheetRetry(Worksheet):

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def __init__(foo, self):
        super().__init__(spreadsheet=self.spreadsheet, properties=self._properties)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def resize(self, **kwargs):
        """resize worksheet to cols, rows count."""
        return super(type(self), self).resize(**kwargs)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def get_values(self, range_name=None, **kwargs):
        """Returns a list of lists containing all cells' values as strings."""
        return super(type(self), self).get_values(range_name, **kwargs)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def get_all_values(self, **kwargs):
        """Returns a list of lists containing all cells' values as strings."""
        return super(type(self), self).get_all_values(**kwargs)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def findall(self, query, in_row=None, in_column=None):
        """Finds all cells matching the query."""
        return super(type(self), self).findall(query=query, in_row=in_row, in_column=in_column)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def col_values(self, col, value_render_option=utils.ValueRenderOption.formatted):
        return super(type(self), self).col_values(col, value_render_option=value_render_option)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def row_values(self, row, value_render_option=utils.ValueRenderOption.formatted):
        return super(type(self), self).row_values(row, value_render_option=value_render_option)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def delete_rows(self, start_index, end_index=None):
        return super(type(self), self).delete_rows(start_index, end_index=end_index)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def update_cell(self, row, col, value):
        return super(type(self), self).update_cell(row=row, col=col, value=value)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def update_cells(self, cell_list, value_input_option=utils.ValueInputOption.raw):
        return super(type(self), self).update_cells(cell_list=cell_list,
                                                    value_input_option=value_input_option)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def batch_clear(self, ranges):
        return super(type(self), self).batch_clear(ranges=ranges)

    @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    def range(self, name):
        return super(type(self), self).range(name=name)

    # """
    # gspread_formatting
    # """
    # @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    # def get_effective_format(self, label):
    #     """Returns a CellFormat object or None representing the effective formatting directives,"""
    #     return functions.get_effective_format(worksheet=self, label=label)

    # @retry(tries=15, delay=2, backoff=2, except_retry=[error_quota_req])
    # def get_user_entered_format(self, label):
    #     """Returns a CellFormat object or None representing the user-entered formatting directives,"""
    #     return functions.get_user_entered_format(worksheet=self, label=label)
