from pygsheets.exceptions import InvalidColorRangeError
from collections import abc
import json


class Color(abc.MutableMapping):

    @classmethod
    def from_json(cls, source):
        if isinstance(source, str):
            source = json.loads(source)
        return cls(**source)

    def __init__(self, red, green, blue, alpha=1):
        if not (self._check_range(red) and self._check_range(green)
                and self._check_range(blue) and self._check_range(alpha)):
            raise InvalidColorRangeError('A color or alpha value is not in range [0, 1]: R:%d; G:%d; B:%d; A:%d.' %
                                         (red, green, blue, alpha))
        self._red = red
        self._green = green
        self._blue = blue
        self._alpha = alpha

    @property
    def red(self):
        return self._red

    @red.setter
    def red(self, value):
        if self._check_range(value):
            self._red = value
        else:
            raise InvalidColorRangeError('Value needs to be in range [0, 1], but is: ' + str(value))

    @property
    def green(self):
        return self._green

    @green.setter
    def green(self, value):
        if self._check_range(value):
            self._green = value
        else:
            raise InvalidColorRangeError('Value needs to be in range [0, 1], but is: ' + str(value))

    @property
    def blue(self):
        return self._blue

    @blue.setter
    def blue(self, value):
        if self._check_range(value):
            self._blue = value
        else:
            raise InvalidColorRangeError('Value needs to be in range [0, 1], but is: ' + str(value))

    @property
    def alpha(self):
        return self._alpha

    @alpha.setter
    def alpha(self, value):
        if self._check_range(value):
            self._alpha = value
        else:
            raise InvalidColorRangeError('Value needs to be in range [0, 1], but is: ' + str(value))

    @staticmethod
    def _check_range(value):
        return 0 <= value <= 1

    def to_json(self):
        return {
            'red': self._red,
            'green': self._green,
            'blue': self._blue,
            'alpha': self._alpha
        }

    def __iter__(self):
        return iter(('red', 'green', 'blue', 'alpha'))

    def __len__(self):
        return 4

    def __setitem__(self, key, value):
        if isinstance(key, str):
            if hasattr(self, key):
                setattr(self, key, value)
            else:
                raise KeyError('Invalid key: %s' % key)
        elif isinstance(key, int):
            setattr(self, ('red', 'green', 'blue', 'alpha')[key], value)

    def __delitem__(self, key):
        if isinstance(key, str):
            if hasattr(self, key):
                setattr(self, key, 0)
            else:
                raise KeyError('Invalid key: %s.' % key)
        elif isinstance(key, int):
            setattr(self, ('red', 'green', 'blue', 'alpha')[key], 0)

    def __getitem__(self, item):
        if isinstance(item, str):
            if hasattr(self, item):
                return getattr(self, item)
            else:
                raise KeyError('Invalid key: %s.' % item)
        elif isinstance(item, int):
            return (self.red, self.green, self.blue, self.alpha)[item]

    def __repr__(self):
        return '<Color R {:.0%}/G {:.0%}/B {:.0%} A {:.0%}>'.format(self._red, self._green, self._blue, self._alpha)

    @classmethod
    def WHITE(cls, alpha=1.0):
        return cls(1, 1, 1, alpha)

    @classmethod
    def NAVY(cls, alpha=1.0):
        return cls(0, 0, 0.5, alpha)

    @classmethod
    def BLACK(cls, alpha=1.0):
        return cls(0, 0, 0, alpha)

    @classmethod
    def RED(cls, alpha=1.0):
        return cls(1, 0, 0, alpha)

    @classmethod
    def YELLOW(cls, alpha=1.0):
        return cls(1, 1, 0, alpha)

    @classmethod
    def GREEN(cls, alpha=1.0):
        return cls(0, 1, 0, alpha)

    @classmethod
    def BLUE(cls, alpha=1.0):
        return cls(0, 0, 1, alpha)


if __name__ == '__main__':
    color = Color(0, 1, 1, 1)

    print(color.to_json())
    color['blue'] = 0.5
    del color['blue']
    print(color['blue'])

    color[2] = 1
    print(color[2])
    print(color)
    print(color[3])

    color_2 = Color.from_json(color.to_json())
    color_2['green'] = 0.45212

    print(color)
    print(color_2)
    print(color_2[1])
    # color = Color(0, 2, -2, 1)
    # color = Color(0, 1, -2, 1)
    color = Color.YELLOW(alpha=0.5)
    print(color)


