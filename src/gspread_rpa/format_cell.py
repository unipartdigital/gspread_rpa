import json
from collections import namedtuple
from math import sqrt

"""
example

    cf01 = CellFormat()
    cf01.background_color('white')

    cf01.top.border_color('black')
    cf01.left.border_style('solid')
    cf01.right.border_style('dashed')
    cf01.bottom.border_style('dashed')

    cf01.horizontal_alignment('center')
    cf01.vertical_alignment('top')
    cf01.wrap_strategy('clip')
    cf01.text_direction("left_to_right")

    cf01.text.foreground_color("red")
    cf01.text.bold(True)
    cf01.text.font_family("Arial")
    cf01.text.font_size(12)
    cf01.text.italic(True)
    cf01.text.strikethrough(False)
    cf01.text.underline(True)
    cf01.text.link(uri="http://example.org")

    cf01.text_rotation.angle(90)
    cf01.text_rotation.vertical(True)

    print (cf01.json_dump())

"""

class ColorMap(object):
    def __init__(self):
        self.white   = {"red": "1.00", "green": "1.00", "blue": "1.00"}
        self.silver  = {"red": "0.75", "green": "0.75", "blue": "0.75"}
        self.gray    = {"red": "0.50", "green": "0.50", "blue": "0.50"}
        self.black   = {"red": "0.00", "green": "0.00", "blue": "0.00"}
        self.red     = {"red": "1.00", "green": "0.00", "blue": "0.00"}
        self.maroon  = {"red": "0.50", "green": "0.00", "blue": "0.00"}
        self.yellow  = {"red": "1.00", "green": "1.00", "blue": "0.00"}
        self.olive   = {"red": "0.50", "green": "0.50", "blue": "0.00"}
        self.lime    = {"red": "0.00", "green": "1.00", "blue": "0.00"}
        self.green   = {"red": "0.00", "green": "0.50", "blue": "0.00"}
        self.aqua    = {"red": "0.00", "green": "1.00", "blue": "1.00"}
        self.teal    = {"red": "0.00", "green": "0.50", "blue": "0.50"}
        self.blue    = {"red": "0.00", "green": "0.00", "blue": "1.00"}
        self.navy    = {"red": "0.00", "green": "0.00", "blue": "0.50"}
        self.fuchsia = {"red": "1.00", "green": "0.00", "blue": "1.00"}
        self.purple  = {"red": "0.50", "green": "0.00", "blue": "0.50"}

        self.names = ['white', 'silver', 'gray', 'black', 'red', 'maroon', 'yellow', 'olive',
                      'lime', 'green', 'aqua', 'teal', 'blue', 'navy', 'fuchsia', 'purple']

    def find_name(self, r,g,b):
        result = []
        RGB = namedtuple('RGB', 'r,g,b')
        e = RGB (float(r), float(g), float(b))
        for c in self.names:
            current = getattr (self, c)
            current = RGB (float(current["red"]), float(current["green"]), float(current["blue"]))
            result.append ((c, sqrt((e.r - current.r)**2 + (e.g - current.g)**2 + (e.b - current.b)**2)))
        result.sort(key=lambda x: x[1])
        return (result[0][0])

    def color(self, name, opacity=1.0):
        result = {}
        if hasattr(self, name):
            result = getattr(self, name)
            result["alpha"] = opacity
        return result

    def upsert(self, name, red, green, blue):
        setattr(self, name.lower(), {"red": "{:.2f}".format(red), "green": "{:.2f}".format(green),
                                     "blue": "{:.2f}".format(blue)})
        if name.lower() not in self.names: self.names.append (name.lower())
        return self

class EnumObject(object):
    def __init__(self):
        pass

    def value(self, name):
        result = ''
        if hasattr(self, name.lower()):
            result = getattr(self, name.lower())
        return result

class Align(EnumObject):
    def __init__(self):
        pass

    def align(self, name):
        return self.value(name)

class AlignHorizontal(Align):
    def __init__(self):
        super(type(self), self).__init__()
        self.left   = "LEFT"
        self.center = "CENTER"
        self.right  = "RIGHT"
class AlignVertical(Align):
    def __init__(self):
        super(type(self), self).__init__()
        self.top    = "TOP"
        self.middle = "MIDDLE"
        self.bottom = "BOTTOM"

class BorderStyle(EnumObject):
    def __init__(self):
        super(type(self), self).__init__()
        self.dotted = "DOTTED" 	        # The border is dotted.
        self.dashed = "DASHED" 	        # The border is dashed.
        self.solid  = "SOLID" 	        # The border is a thin solid line.
        self.solid_medium = "SOLID_MEDIUM" 	# The border is a medium solid line.
        self.solid_thick = "SOLID_THICK" 	# The border is a thick solid line.
        self.null = "NONE" 	        # No border. Used only when updating a border in order to erase it.
        self.double = "DOUBLE" 	        # The border is two solid lines.

    def style(self, name):
        return self.value(name)

class DictObject(dict):
    def __init__(self):
        pass

    def o2dict(self):
        result = {}
        for fk in self.keys:
            k = fk[0]
            l = fk[1].split(',')
            if hasattr(self, k):
                tmp = getattr (self, k)
                if hasattr (tmp, 'o2dict'):
                    if tmp.o2dict(): tmp = tmp.o2dict()
                c = result
                for i in l:
                    if not tmp: continue
                    if i == l[-1]:
                        c[i] = tmp
                    elif i in c:
                        c = c[i]
                    else:
                        c[i] = {}
                        c = c[i]
        return result

class Border(BorderStyle, ColorMap, DictObject):

    keys = [('_style', 'style'), ('_color', 'color')]

    def __init__(self):
        pass

    def border_style(self, name):
        setattr(self, '_style', BorderStyle().style(name=name))

    def style(self, name):
        setattr(self, '_style', BorderStyle().style(name=name))

    def border_color(self, name, opacity=1.0):
        setattr(self, '_color', ColorMap().color(name=name, opacity=opacity))

    def color(self, rgba):
        setattr(self, '_color', rgba)

""" https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#WrapStrategy """
class WrapStrategy(EnumObject):
    def __init__(self):
        super(type(self), self).__init__()
        self.overflow_cell = "OVERFLOW_CELL"
        self.legacy_wrap   = "LEGACY_WRAP"
        self.clip          = "CLIP"
        self.wrap          = "WRAP"

    def wrap_strategy (self, name):
        return self.value(name)
    wrapStrategy = wrap_strategy

class TextDirection(EnumObject):
    def __init__(self):
        super(type(self), self).__init__()
        self.left_to_right  = "LEFT_TO_RIGHT"
        self.right_to_left  = "RIGHT_TO_LEFT"

    def text_direction(self, name):
        return self.value(name)
    textDirection = text_direction

class Link(DictObject):

    keys = [('_uri', 'uri')]
    def __init__(self, uri=''):
        if uri: self._uri = uri["uri"] if "uri" in uri else uri

    def uri(self, uri):
        uri = uri["uri"] if "uri" in uri else uri
        setattr(self, '_uri', uri)

class TextFormat(DictObject):

    keys = [('_foregroundColor','foregroundColor'), ('_fontFamily','fontFamily'), ('_fontSize','fontSize'),
            ('_bold', 'bold'), ('_italic','italic'), ('_strikethrough','strikethrough'),
            ('_strikethrough','strikethrough'), ('_underline','underline'), ('_link', 'link')]

    def __init__(self):
        pass

    def foreground_color(self, name, opacity=1.0):
        setattr(self, '_foregroundColor', ColorMap().color(name=name, opacity=opacity))

    def foregroundColor(self, rgba):
        setattr(self, '_foregroundColor', rgba)

    def font_family(self, name):
        setattr(self, '_fontFamily', name)
    fontFamily = font_family

    def font_size(self, size):
        assert int(size)
        setattr(self, '_fontSize', size)
    fontSize = font_size

    def bold(self, value):
        setattr(self, '_bold', value)

    def italic(self, value):
        setattr(self, '_italic', value)

    def strikethrough(self, value):
        setattr(self, '_strikethrough', value)

    def underline (self, value):
        setattr(self, '_underline', value)

    def link(self, uri):
        setattr(self, '_link', Link(uri))

class TextRotation(DictObject):

    keys = [('_angle', 'angle'), ('_vertical', 'vertical')]
    def __init__(self, angle=None, vertical=None):
        pass

    def angle(self, angle):
        setattr(self, '_angle', angle)

    def vertical(self, vertical):
        setattr(self, '_vertical', vertical)

class CellFormat(Border, AlignVertical, AlignHorizontal, TextFormat, DictObject):

    keys = [('_backgroundColor', 'backgroundColor'),
            ('top','borders,top') , ('bottom','borders,bottom'),
            ('left', 'borders,left'), ('right', 'borders,right'),
            ('_horizontalAlignment','horizontalAlignment'), ('_verticalAlignment', 'verticalAlignment'),
            ('_wrapStrategy','wrapStrategy'), ('_textDirection','textDirection'),
            ('text', 'textFormat'), ('text_rotation', 'textRotation')]

    def __init__(self, other=None):
        if other:
            self = self.dict2o(other.o2dict())
        else:
            self.top = Border()
            self.bottom = Border()
            self.left = Border()
            self.right = Border()
            self.text = TextFormat()
            self.text_rotation = TextRotation()

    @property
    def textFormat (self):
        return self.text

    @property
    def textRotation(self):
        return self.text_rotation

    @property
    def background_color_name (self):
        if hasattr (self, '_backgroundColor'):
            tmp = getattr (self, '_backgroundColor')
            r = tmp['red']   if 'red'   in tmp else 0.0
            g = tmp['green'] if 'green' in tmp else 0.0
            b = tmp['blue']  if 'blue'  in tmp else 0.0
            return ColorMap().find_name (r, g, b)
        return ''

    def background_color (self, name, opacity=1.0):
        setattr(self, '_backgroundColor', ColorMap().color(name=name, opacity=opacity))

    def backgroundColor (self, rgba):
        setattr(self, '_backgroundColor', rgba)

    def horizontal_alignment(self, name):
        if name: setattr(self, '_horizontalAlignment', AlignHorizontal().align(name=name))
    horizontalAlignment = horizontal_alignment

    def vertical_alignment(self, name):
        if name: setattr(self, '_verticalAlignment', AlignVertical().align(name=name))
    verticalAlignment = vertical_alignment

    def wrap_strategy(self, name):
        if name: setattr(self, '_wrapStrategy', WrapStrategy().wrap_strategy(name=name))
    wrapStrategy = wrap_strategy

    def text_direction(self, name):
        if name: setattr(self, '_textDirection', TextDirection().text_direction(name=name))
    textDirection = text_direction

    def json_dump(self):
        result = json.dumps(self.o2dict(), indent=4, sort_keys=True)
        return result

    def dict2o(self, d):
        def apply (this, e):
            for k, v in e.items():
                # print ("k: {} v: {}".format(k, v))
                if hasattr(this, k):
                    # print ("hasattr {}".format(k))
                    tmp = getattr(this, k)
                    if hasattr(tmp, '__call__'):
                        # print ("call {} {}".format (tmp, v))
                        tmp(v)
                    elif hasattr(tmp, '__dict__') and isinstance (v, dict):
                        # print ("__dict__ {}".format (v))
                        apply(tmp, v)
                elif isinstance (e[k], dict):
                    apply (this, e[k])
        result = CellFormat()
        """ flatten 'boders: {top, left ...}' into top, left ..."""
        if 'borders' in d:
            borders = d.pop('borders')
            d.update (borders)
        apply(result, d)
        return result

if __name__ == "__main__":

    cf01 = CellFormat()
    cf01.background_color('white')

    cf01.top.border_color('black')
    cf01.left.border_style('solid')
    cf01.right.border_style('dashed')
    cf01.bottom.border_style('dashed')

    cf01.horizontal_alignment('center')
    cf01.vertical_alignment('top')
    cf01.wrap_strategy('clip')
    cf01.text_direction("left_to_right")

    cf01.text.foreground_color("red")
    cf01.text.bold(True)
    cf01.text.font_family("Arial")
    cf01.text.font_size(12)
    cf01.text.italic(True)
    cf01.text.strikethrough(False)
    cf01.text.underline(True)
    cf01.text.link(uri="http://example.org")

    cf01.text_rotation.angle(90)
    cf01.text_rotation.vertical(True)

    j = json.loads(cf01.json_dump())
    new_o = cf01.dict2o(j)

    assert cf01.json_dump() != "{}"
    assert cf01.json_dump() == new_o.json_dump()
    assert CellFormat().json_dump() == "{}"
    assert cf01.background_color_name == new_o.background_color_name
