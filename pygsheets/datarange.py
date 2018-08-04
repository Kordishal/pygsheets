# -*- coding: utf-8 -*-.

"""
pygsheets.datarange
~~~~~~~~~~~~~~~~~~~

This module contains DataRange class for storing/manipulating a range of data in spreadsheet. This class can
be used for group operations, e.g. changing format of all cells in a given range. This can also represent named ranges
protected ranges, banned ranges etc.

"""
from pygsheets.exceptions import InvalidArgumentValue, CellNotFound, IncorrectCellLabel, InvalidRange
from pygsheets.custom_types import DateTimeRenderOption, ValueRenderOption, Dimension, ValueInputOption

from collections.abc import Sequence
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
                raise InvalidArgumentValue('Address coordinates may not be below zero: ' + repr(value))
            self._value = value
        elif isinstance(value, Address):
            self._value = self._label_to_coordinates(value.label)
        else:
            raise IncorrectCellLabel('Only labels in A1 notation, coordinates as a tuple or '
                                     'pygsheets.Address objects are accepted.')

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

    def __repr__(self):
        return self.label

    def __iter__(self):
        return iter(self._value)

    def __getitem__(self, item):
        return self._value[item]

    def __eq__(self, other):
        if isinstance(other, Address):
            return self.label == other.label
        else:
            return NotImplemented

    def __ne__(self, other):
        result = self.__eq__(other)
        if result is NotImplemented:
            return result
        return not result


class Range(object):
    """A range in a google sheet.

    :param start:   The start of the range (either as tuple, A1 notation or pygsheets.Address object.)
    :param end:     The end of the range (either as tuple, A1 notation or pygsheets.Address object.)
    """

    def __init__(self, start, end):
        self._start_address = Address(start)
        self._end_address = Address(end)
        if self._start_address[0] > self._end_address[0] or self._start_address[1] > self._end_address[1]:
            raise InvalidRange('Cannot define a range with the start address after the end address: {}'.format(self.range))

    @property
    def start(self):
        """Top left address of this range."""
        return self._start_address

    @property
    def end(self):
        """Bottom right address of this range."""
        return self._end_address

    @property
    def range(self):
        """The range in A1 notation: 'A1:B7'"""
        return self.__repr__()

    def __repr__(self):
        return '{}:{}'.format(self.start.label, self.end.label)

    def __eq__(self, other):
        if isinstance(other, Range):
            return self.range == other.range
        else:
            return NotImplemented

    def __ne__(self, other):
        result = self.__eq__(other)
        if result is NotImplemented:
            return result
        return not result


class GridRange(Range):

    def __init__(self, worksheet, start, end):
        super().__init__(start, end)
        self._sheet = worksheet

    @property
    def worksheet_id(self):
        """The id of the worksheet this grid range is part of."""
        return self._sheet.id

    @property
    def worksheet(self):
        """The worksheet this grid range is part of."""
        return self._sheet

    @property
    def grid_range(self):
        """The grid range as Google Sheets API JSON object."""
        return self.to_json()

    def to_json(self):
        """The grid range as Google Sheets API JSON object."""
        return {
            'sheetId': self.worksheet_id,
            'startRowIndex': self.start[0],
            'endRowIndex': self.end[0] + 1,
            'startColumnIndex': self.start[1],
            'endColumnIndex': self.end[1] + 1
        }

    def __repr__(self):
        """The grid range in A1 notation."""
        return '{}!{}:{}'.format(self.worksheet.title, self.start.label, self.end.label)


class ValueRange(GridRange, Sequence):

    def __init__(self, worksheet, start, end, major_dimension=Dimension.ROWS):
        super().__init__(worksheet, start, end)
        self._major_dimension = major_dimension
        self._value_render_option = ValueRenderOption.FORMATTED_VALUE
        self._date_time_render_option = DateTimeRenderOption.FORMATTED_STRING
        self._values = None
        self.load()

    @property
    def major_dimension(self):
        """The major dimension of this value range. When changed, the values are reordered accordingly."""
        return self._major_dimension

    @major_dimension.setter
    def major_dimension(self, value):
        if isinstance(value, Dimension):
            value = value.value
        if self._major_dimension.name != value:
            self._major_dimension = Dimension[value]
            new_values = list()
            for i in range(len(self._values[0])):
                dimension = list()
                for item in self._values:
                    dimension.append(item[i])
                new_values.append(dimension)
            self._values = new_values

    @property
    def value_render_option(self):
        return self._value_render_option

    @value_render_option.setter
    def value_render_option(self, value):
        self._value_render_option = value

    @property
    def date_time_render_option(self):
        return self._date_time_render_option

    @date_time_render_option.setter
    def date_time_render_option(self, value):
        self._date_time_render_option = value

    def to_json(self):
        return {
            'range': self.range,
            'majorDimension': self.major_dimension.value,
            'values': self._values
        }

    def save(self, value_input_option=ValueInputOption.USER_ENTERED, include_values_in_response=False):
        """Saves all values in the spreadsheet."""
        return self.worksheet.client.sheet.values_update(self.worksheet.spreadsheet.id,
                                                         self.range,
                                                         self.to_json(),
                                                         value_input_option,
                                                         include_values_in_response,
                                                         response_value_render_option=self.value_render_option,
                                                         response_date_time_render_option=self.date_time_render_option)

    def load(self):
        """Loads all values from spreadsheet. Careful: this will overwrite unsaved changes."""
        response = self.worksheet.client.sheet.values_get(self.worksheet.spreadsheet.id,
                                                          self.range,
                                                          self.major_dimension,
                                                          self.value_render_option,
                                                          self.date_time_render_option)
        self._values = response['values']

    def __eq__(self, other):
        if isinstance(other, ValueRange):
            return self.range == other.range and self.worksheet.spreadsheet.id == other.worksheet.spreadsheet.id
        else:
            return NotImplemented

    def __ne__(self, other):
        result = self.__eq__(other)
        if result is NotImplemented:
            return result
        return not result

    def __iter__(self):
        return self._values

    def __len__(self):
        return sum([len(v) for v in self._values])

    def __getitem__(self, item):
        return self._values[item]

    def __repr__(self):
        return '<{} "{}">'.format(self.__class__.__name__, self.range)


class DataRange(GridRange):
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
        return '{}:{}'.format(self._start_address, self._end_address)

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


class NamedRange(DataRange):

    def __init__(self, name, worksheet, start, end):
        super().__init__(worksheet, start, end)
        self._name = name
        self._id = None
        self.load()

    @property
    def id(self):
        return self._id

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        self._name = value

    def to_json(self):
        return {
            "namedRangeId": self._id,
            "name": self._name,
            "range": self.grid_range
        }

    def save(self):
        pass

    def load(self):
        pass


class ProtectedRange(DataRange):

    def __init__(self):
        self._protected_id = None
        self.description = ''
        self.warningOnly = False
        self.requestingUserCanEdit = False
        self.editors = None
