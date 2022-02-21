# gspread_rpa

 a [gspread](https://docs.gspread.org/en/latest/) (Python API for Google Sheets) hight level wrapper

## Object:

* CellIndex, GridIndex
  * use x,y (col, row) indexes start at 1
* GoogleSheets
  * embedded spreadsheets and worksheets as well as revision download upload
* CellFormat
  * function call cells formating

## Usage Examples

```
python3 -m pip install --user gspread-rpa
```

```
from gspread_rpa import CellIndex, GridIndex, GoogleSheets, CellFormat

```

 [Basic usage examples](src/gspread_rpa/demo/)


## Contribute and contact

* via the [github repository](https://github.com/unipartdigital/gspread_rpa)
* by email: <rpa@unipart.io>

## License notices

```
This program is free software: you can redistribute it and/or modify it under the terms of the GNU
General Public License as published by the Free Software Foundation, either version 3 of the License,
or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.
If not, see <https://www.gnu.org/licenses/>.
```
