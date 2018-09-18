# -*- coding: utf-8 -*-.

"""
pygsheets.ranges
~~~~~~~~~~~~~~~~~~~

A collection of support classes to deal with various types of ranges and addresses of the Google Sheets API.
"""

from pygsheets.exceptions import InvalidArgumentValue, IncorrectCellLabel
from pygsheets.custom_types import ValueRenderOption, DateTimeRenderOption, Dimension, ValueInputOption

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
    >>> str(a)
    A1
    """

    _MAGIC_NUMBER = 64

    def __init__(self, value):
        if isinstance(value, str):
            self._value = self._label_to_coordinates(value)
        elif isinstance(value, tuple):
            if value[0] < 1 or value[1] < 1:
                raise InvalidArgumentValue('Coordinates in a sheet cannot be below zero: ' + repr(value))
            self._value = value
        elif isinstance(value, Address):
            self._value = self._label_to_coordinates(value.label)
        else:
            raise IncorrectCellLabel('Valid cell labels include the A1 notation, tuple coordinates and '
                                     'pygsheets.Address objects.')

    @property
    def label(self):
        """The label of this address in A1 notation."""
        return self._value_as_label()

    def _value_as_label(self):
        """Transforms tuple coordinates into a label of the form A1."""
        row = int(self._value[0])
        if row < 1:
            raise InvalidArgumentValue('Coordinates in a sheet cannot be below zero: ' + repr(self._value))
        row_label = str(row)

        col = int(self._value[1])
        if col < 1:
            raise InvalidArgumentValue('Coordinates in a sheet cannot be below zero: ' + repr(self._value))
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

    def __str__(self):
        return self.label

    def __repr__(self):
        return self.label

    def __iter__(self):
        return iter(self._value)

    def __getitem__(self, item):
        return self._value[item]

    def __eq__(self, other):
        if isinstance(other, Address):
            return self.label == other.label
        elif isinstance(other, str):
            return self.label == other
        elif isinstance(other, tuple):
            return self._value == other
        else:
            return NotImplemented

    def __ne__(self, other):
        result = self.__eq__(other)
        if result is NotImplemented:
            return result
        return not result



class ValueRange(object):

    def __init__(self, worksheet, start, end, major_dimension=Dimension.ROWS):
        super().__init__(worksheet, start, end)
        self._start = Address(start)
        self._end = Address(end)
        self._worksheet = worksheet
        self._major_dimension = major_dimension
        self._value_render_option = ValueRenderOption.FORMATTED_VALUE
        self._date_time_render_option = DateTimeRenderOption.FORMATTED_STRING
        self._values = None
        self.load()

    @property
    def start(self):
        return self._start

    @start.setter
    def start(self, value):
        self._start = Address(value)

    @property
    def end(self):
        return self._end

    @end.setter
    def end(self, value):
        self._end = Address(value)

    @property
    def worksheet(self):
        return self._worksheet

    @property
    def range(self):
        return '{}:{}'.format(self.start, self.end)

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
        if isinstance(value, ValueRenderOption):
            self._value_render_option = value
        else:
            self._value_render_option = ValueRenderOption[value]

    @property
    def date_time_render_option(self):
        return self._date_time_render_option

    @date_time_render_option.setter
    def date_time_render_option(self, value):
        if isinstance(value, DateTimeRenderOption):
            self._date_time_render_option = value
        else:
            self._date_time_render_option = DateTimeRenderOption[value]

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



