import pytest

from pygsheets.ranges import Address
from pygsheets.exceptions import IncorrectCellLabel, InvalidArgumentValue


class TestAddress(object):

    def setup_class(self):
        self.a = Address('A1')
        self.b = Address('A1')

    def test_init_input_label(self):
        assert isinstance(Address('A1'), Address)

    def test_init_with_coordinate(self):
        assert isinstance(Address((25, 5)), Address)

    def test_init_with_address(self):
        assert isinstance(Address(Address('AB92')), Address)

    def test_init_with_invalid_label_input(self):
        with pytest.raises(IncorrectCellLabel):
            Address('AV')

        with pytest.raises(IncorrectCellLabel):
            Address('32AX')

    def test_init_with_invalid_coordinates(self):
        with pytest.raises(InvalidArgumentValue):
            Address((-1, 0))

    def test_label(self):
        assert self.a.label == 'A1'

    def test_getitem(self):
        assert self.a[0] == 1
        assert self.a[1] == 1

    def test_iter(self):
        for i in self.a:
            assert i == 1

    def test_repr(self):
        assert str(self.a) == 'A1'

    def test_eq(self):
        assert self.a == self.b