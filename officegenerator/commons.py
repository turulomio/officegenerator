## @namespace officegenerator.commons
## @brief Common code to odfpy and openpyxl wrappers
import datetime
import functools
import gettext
import os
import pkg_resources
import warnings
from odf.opendocument import  __version__ as __odfpy_version__

__version__ = '0.9.0'
__versiondate__=datetime.date(2018,9,6)

try:
    t=gettext.translation('officegenerator',pkg_resources.resource_filename("officegenerator","locale"))
    _=t.gettext
except:
    _=str


def deprecated(func):
     """This is a decorator which can be used to mark functions
     as deprecated. It will result in a warning being emitted
     when the function is used."""
     @functools.wraps(func)
     def new_func(*args, **kwargs):
         warnings.simplefilter('always', DeprecationWarning)  # turn off filter
         warnings.warn("Call to deprecated function {}.".format(func.__name__),
                       category=DeprecationWarning,
                       stacklevel=2)
         warnings.simplefilter('default', DeprecationWarning)  # reset filter
         return func(*args, **kwargs)
     return new_func


## Function used in argparse_epilog
## @return String
def argparse_epilog():
    return _("Developed by Mariano Muñoz 2015-{}").format(__versiondate__.year)


## Allows to operate with columns letter names
## @param letter String with the column name. For example A or AA...
## @param number Columns to move
## @return String With the name of the column after movement
def columnAdd(letter, number):
    letter_value=column2number(letter)+number
    return number2column(letter_value)


def rowAdd(letter,number):
    return str(int(letter)+number)

## Convierte un número  con el numero de columna al nombre de la columna de hoja de datos
##
## Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA.
def number2column(n):
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

## Convierte una columna de hoja de datos a un número
##
## Excel-style column name to number, e.g., A = 1, Z = 26, AA = 27, AAA = 703.
def column2number(name):
    n = 0
    for c in name:
        n = n * 26 + 1 + ord(c) - ord('A')
    return n

## Converts a column name to a index position (number of column -1)
def column2index(name):
    return column2number(name)-1

## Convierte el nombre de la fila de la hoja de datos a un índice, es decir el número de la fila -1
def row2index(number):
    return int(number)-1

## Covierte el nombre de la fila de la hoja de datos a un  numero entero que corresponde con el numero de la fila
def row2number(strnumber):
    return int(strnumber)

## Convierte el numero de la fila al nombre de la fila en la hoja de datos , que corresponde con un string del numero de la fila
def number2row(number):
    return str(number)
    
## Convierte el indice de la fila al numero cadena de la hoja de datos
def index2row(index):
    return str(index+1)
    
## Convierte el indice de la columna a la cadena de letras de la columna de la hoja de datos
def index2column(index):
    return number2column(index+1)
    
## Crea un directorio con todos sus subdirectorios
##
## No produce error si ya está creado.
def makedirs(dir):
    try:
        os.makedirs(dir)
    except:
        pass

def ODFPYversion():
    return __odfpy_version__.split("/")[1]


class Coord:
    def __init__(self, strcoord):
        self.letter, self.number=self.__extract(strcoord)
    def __extract(self, strcoord):
        if strcoord.find(":")!=-1:
            print("I can't manage range coord")
            return
        letter=""
        number=""
        for l in strcoord:
            if l.isdigit()==False:
                letter=letter+l
            else:
                number=number+l
        return (letter,number)

    def addRow(self, num=1):
        self.number=str(int(self.number)+num)
        return self

    @staticmethod
    def assertCoord(o):
        if o.__class__==Coord:
            return o
        elif o.__class__==str:
            return Coord(o)

class Range:
    def __init__(self,strrange):
        self.start, self.end=self.__extract(strrange)

    def __extract(self,range):
        if range.find(":")==-1:
            print("I can't manage this range")
            return
        a=range.split(":")
        return (Coord(a[0]), Coord(a[1]))

    def numRows(self):
        return row2number(self.start.number)-row2number(self.end.number)

    def numColumns(self):
        return column2number(self.start.letter)-column2number(self.end.letter)

    @staticmethod
    def assertRange(o):
        if o.__class__==Range:
            return o
        elif o.__class__==str:
            return Range(o)
