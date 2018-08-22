# -*- coding: utf-8 -*-.

"""
pygsheets.ranges
~~~~~~~~~~~~~~~~~~~

A collection of support classes to deal with various types of ranges and addresses of the Google Sheets API.
"""

from pygsheets.exceptions import InvalidArgumentValue, IncorrectCellLabel

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

