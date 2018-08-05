import pytest

import pygsheets
from pygsheets.worksheet import Worksheet
from pygsheets.datarange import GridRange


class TestGridRange(object):

    def setup_class(self):
        c = pygsheets.authorize()
        ss = c.create('grid_range_test_sheet')
        self.wks = ss.sheet1
        self.grid_range = GridRange(self.wks, 'A5', 'B10')

    def test_init(self):
        assert isinstance(GridRange(self.wks, 'A1', 'B5'), GridRange)

    def test_worksheet_id(self):
        assert self.grid_range.worksheet_id == self.wks.id

    def test_worksheet(self):
        assert isinstance(self.grid_range.worksheet, Worksheet)

    def test_grid_range(self):
        assert self.grid_range.grid_range == {
            'sheetId': self.wks.id,
            'startRowIndex': 5,
            'endRowIndex': 11,
            'startColumnIndex': 1,
            'endColumnIndex': 3}

    def test_to_json(self):
        assert self.grid_range.to_json() == {
            'sheetId': self.wks.id,
            'startRowIndex': 5,
            'endRowIndex': 11,
            'startColumnIndex': 1,
            'endColumnIndex': 3}

    def test_repr(self):
        assert self.grid_range.__repr__() == 'Sheet1!A5:B10'


class TestNamedRanges(object):

    def setup_class(self):
        c = pygsheets.authorize()
        ss = c.create('grid_range_test_sheet')
        self.wks = ss.sheet1
        self.wks.add_named_range('test_range', 'A1', 'B7')
        self.named_range = self.wks.spreadsheet.named_ranges['test_range']

    def test_name(self):
        assert self.named_range == 'test_range'