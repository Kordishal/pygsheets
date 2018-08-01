from pygsheets.datarange import Address


class TestAddress(object):

    def test_constructor(self):
        a = Address('A1')
        assert a.label == 'A1'
        assert a[0] == 1
        assert a[1] == 1

        b = Address('E25')
        assert b.label == 'E25'
        assert b[0] == 25
        assert b[1] == 5
