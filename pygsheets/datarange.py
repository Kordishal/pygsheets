# -*- coding: utf-8 -*-.

"""
pygsheets.datarange
~~~~~~~~~~~~~~~~~~~

This module contains DataRange class for storing/manipulating a range of data in spreadsheet. This class can
be used for group operations, e.g. changing format of all cells in a given range. This can also represent named ranges
protected ranges, banned ranges etc.

"""

from pygsheets.utils import format_addr
from pygsheets.exceptions import InvalidArgumentValue, CellNotFound, IncorrectCellLabel

import warnings
import re


class Address(object):
    """Represents the address of a cell.

    >>> a = Address('A1')
    >>> a.label
    A1
    >>> a[0]
    1
    >>> a[1]
    1
    >>> a = Address((1, 1))
    >>> a.label
    A1
    """

    _MAGIC_NUMBER = 64

    def __init__(self, value):
        if isinstance(value, str):
            self._value = self._label_to_coordinates(value)
        elif isinstance(value, tuple):
            if value[0] < 1 or value[1] < 1:
                raise InvalidArgumentValue('Address coordinates may not be below zero: ' + repr(self._value))
            self._value = value
        elif isinstance(value, Address):
            self._value = self._label_to_coordinates(value.label)
        else:
            raise IncorrectCellLabel('Only labels in A1 notation or coordinates as a tuple are accepted.')

    @property
    def label(self):
        return self._value_as_label()

    def _value_as_label(self):
        """Transforms tuple coordinates into a label of the form A1."""
        row = int(self._value[0])
        if row < 1:
            raise InvalidArgumentValue('Address coordinates may not be below zero: ' + repr(self._value))
        row_label = str(row)

        col = int(self._value[1])
        if col < 1:
            raise InvalidArgumentValue('Address coordinates may not be below zero: ' + repr(self._value))
        div = col
        column_label = ''
        while div:
            (div, mod) = divmod(div, 26)
            if mod == 0:
                mod = 26
                div -= 1
            column_label = chr(mod + self._MAGIC_NUMBER) + column_label
        return '{}{}'.format(column_label, row_label)

    def _label_to_coordinates(self, label):
        """Transforms a label in A1 notation into numeric coordinates and returns them as tuple."""
        m = re.match(r'([A-Za-z]+)(\d+)', label)
        if m:
            column_label = m.group(1).upper()
            row, col = int(m.group(2)), 0
            for i, c in enumerate(reversed(column_label)):
                col += (ord(c) - self._MAGIC_NUMBER) * (26 ** i)
        else:
            raise IncorrectCellLabel('Not a valid cell label format: {}.'.format(label))
        return int(row), int(col)

    def __iter__(self):
        return self._value

    def __getitem__(self, item):
        return self._value[item]

    def __eq__(self, other):
        return self.label == other.label


class DataRange(object):
    """
    DataRange specifies a range of cells in the sheet

    :param start: top left cell address
    :param end: bottom right cell address
    :param worksheet: worksheet where this range belongs
    :param name: name of the named range
    :param data: data of the range in as row major matrix
    :param name_id: id of named range
    :param namedjson: json representing the NamedRange from api
    """

    def __init__(self, start=None, end=None, worksheet=None, name='', data=None, name_id=None, namedjson=None, protect_id=None, protectedjson=None):
        self._worksheet = worksheet
        if namedjson:
            start = (namedjson['range'].get('startRowIndex', 0)+1, namedjson['range'].get('startColumnIndex', 0)+1)
            # @TODO this won't scale if the sheet size is changed
            end = (namedjson['range'].get('endRowIndex', self._worksheet.cols),
                   namedjson['range'].get('endColumnIndex', self._worksheet.rows))
            name_id = namedjson['namedRangeId']
        if protectedjson:
            start = (protectedjson['range'].get('startRowIndex', 0)+1, protectedjson['range'].get('startColumnIndex', 0)+1)
            # @TODO this won't scale if the sheet size is changed
            end = (protectedjson['range'].get('endRowIndex', self._worksheet.cols),
                   protectedjson['range'].get('endColumnIndex', self._worksheet.rows))
            protect_id = protectedjson['protectedRangeId']
        self._start_address = Address(start)
        self._end_address = Address(end)
        if data:
            if len(data) == self._end_address[0] - self._start_address[0] + 1 and \
                            len(data[0]) == self._end_address[1] - self._start_address[1] + 1:
                self._data = data
            else:
                self.fetch()
        else:
            self.fetch()

        self._linked = True

        self._name_id = name_id
        self._protect_id = protect_id
        self._name = name

        self.protected_properties = ProtectedRange()
        self._banned = False

    @property
    def name(self):
        """name of the named range. setting a name will make this a range a named range
            setting this to empty string will delete the named range
        """
        return self._name

    @name.setter
    def name(self, name):
        if type(name) is not str:
            raise InvalidArgumentValue('name should be a string')
        if name == '':
            self._worksheet.delete_named_range(self._name)
            self._name = ''
        else:
            if self._name == '':
                # @TODO handle when not linked (create an range on link)
                self._worksheet.create_named_range(name, start=self._start_address, end=self._end_address)
                self._name = name
            else:
                self._name = name
                if self._linked:
                    self.update_named_range()

    @property
    def name_id(self):
        return self._name_id

    @property
    def protect_id(self):
        return self._protect_id

    @property
    def protected(self):
        """get/set range protection"""
        return self._protect_id is not None

    @protected.setter
    def protected(self, value):
        if value:
            resp = self._worksheet.create_protected_range(self._get_gridrange())
            self._protect_id = resp['replies'][0]['addProtectedRange']['protectedRange']['protectedRangeId']
        elif self._protect_id is not None:
            self._worksheet.remove_protected_range(self._protect_id)
            self._protect_id = None

    @property
    def start_addr(self):
        """top-left address of the range"""
        return self._start_address

    @start_addr.setter
    def start_addr(self, addr):
        self._start_address = Address(addr)
        if self._linked:
            self.update_named_range()

    @property
    def end_addr(self):
        """bottom-right address of the range"""
        return self._end_address

    @end_addr.setter
    def end_addr(self, addr):
        self._end_address = Address(addr)
        if self._linked:
            self.update_named_range()

    @property
    def range(self):
        """Range in format A1:C5"""
        return '{}:{}'.format(self._start_address.label, self._end_address.label)

    @property
    def worksheet(self):
        return self._worksheet

    @property
    def cells(self):
        """Get cells of this range"""
        if len(self._data[0]) == 0:
            self.fetch()
        return self._data

    def link(self, update=True):
        """link the datarange so that all properties are synced right after setting them

        :param update: if the range should be synced to cloud on link
        """
        self._linked = True
        if update:
            self.update_named_range()
            self.update_values()

    def unlink(self):
        """unlink the sheet so that all properties are not synced as it is changed"""
        self._linked = False

    def fetch(self, only_data=True):
        """
        update the range data/properties from cloud

        :param only_data: fetch only data

        """
        self._data = self._worksheet.get_values(self._start_address, self._end_address, returnas='cells',
                                                include_tailing_empty_rows=True)
        if not only_data:
            pass

    def apply_format(self, cell):
        """
        Change format of all cells in the range

        :param cell: a model :class: Cell whose format will be applied to all cells

        """
        request = {"repeatCell": {
            "range": self._get_gridrange(),
            "cell": cell.get_json(),
            "fields": "userEnteredFormat,hyperlink,note,textFormatRuns,dataValidation,pivotTable"
            }
        }
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def update_values(self, values=None):
        """
        Update the values of the cells in this range

        :param values: values as matrix

        """
        if values and self._linked:
            self._worksheet.update_values(crange=self.range, values=values)
            self.fetch()
        if self._linked and not values:
            self._worksheet.update_values(cell_list=self._data)

    # @TODO
    def sort(self):
        warnings.warn('Functionality not implemented')

    def update_named_range(self):
        """update the named properties"""
        if self._name_id == '':
            return False
        request = {'updateNamedRange':{
          "namedRange": {
              "namedRangeId": self._name_id,
              "name": self._name,
              "range": self._get_gridrange(),
          },
          "fields": '*',
        }}
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def _get_gridrange(self):
        return {
            "sheetId": self._worksheet.id,
            "startRowIndex": self._start_address[0] - 1,
            "endRowIndex": self._end_address[0],
            "startColumnIndex": self._start_address[1] - 1,
            "endColumnIndex": self._end_address[1],
        }

    def __getitem__(self, item):
        if len(self._data[0]) == 0:
            self.fetch()
        if type(item) == int:
            try:
                return self._data[item]
            except IndexError:
                raise CellNotFound

    def __eq__(self, other):
        return self.start_addr == other.start_addr and self.end_addr == other.end_addr and self.name == other.name

    def __repr__(self):
        range_str = self.range
        if self.worksheet:
            range_str = str(self.range)
        protected_str = " protected" if hasattr(self, '_protected') and self._protected else ""

        return '<%s %s %s%s>' % (self.__class__.__name__, str(self._name), range_str, protected_str)


class ProtectedRange(object):

    def __init__(self):
        self._protected_id = None
        self.description = ''
        self.warningOnly = False
        self.requestingUserCanEdit = False
        self.editors = None
