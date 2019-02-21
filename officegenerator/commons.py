## @namespace officegenerator.commons
## @brief Common code to odfpy and openpyxl wrappers
import datetime
import functools
import gettext
import logging
import os
import pkg_resources
import sys
import warnings
from decimal import Decimal
from odf.opendocument import  __version__ as __odfpy_version__

__version__ = '1.7.0'
__versiondate__=datetime.date(2019,2,21)

try:
    t=gettext.translation('officegenerator',pkg_resources.resource_filename("officegenerator","locale"))
    _=t.gettext
except:
    _=str


## This is a decorator which can be used to mark functions as deprecated. It will result in a warning being emitted when the function is used
def deprecated(func):
     @functools.wraps(func)
     def new_func(*args, **kwargs):
         warnings.simplefilter('always', DeprecationWarning)  # turn off filter
         warnings.warn("Call to deprecated function {}.".format(func.__name__), category=DeprecationWarning, stacklevel=2)
         warnings.simplefilter('default', DeprecationWarning)  # reset filter
         return func(*args, **kwargs)
     return new_func

## Class to manage currencies in officegenerator
##
## The symbol is defined by code with self.symbol()
class Currency:
    def __init__(self, amount=None,  currency='EUR') :
        if amount==None:
            self.amount=Decimal(0)
        else:
            self.amount=Decimal(str(amount))
        if currency==None:
            self.currency='EUR'
        else:
            self.currency=currency


    def __add__(self, money):
        """Si las divisas son distintas, queda el resultado con la divisa del primero"""
        if self.currency==money.currency:
            return Currency(self.amount+money.amount, self.currency)
        else:
            logging.error("Before adding, please convert to the same currency")
            raise "OdfMoneyOperationException"

    def __sub__(self, money):
        """Si las divisas son distintas, queda el resultado con la divisa del primero"""
        if self.currency==money.currency:
            return Currency(self.amount-money.amount, self.currency)
        else:
            logging.error("Before substracting, please convert to the same currency")
            raise "CurrencyOperationException"

    def __lt__(self, money):
        if self.currency==money.currency:
            if self.amount < money.amount:
                return True
            return False
        else:
            logging.error("Before lt ordering, please convert to the same currency")
            sys.exit(1)

    ## Si las divisas son distintas, queda el resultado con la divisa del primero
    ##
    ## En caso de querer multiplicar por un numero debe ser despues. For example: money*4
    def __mul__(self, money):
        if money.__class__ in (int,  float, Decimal):
            return Currency(self.amount*money, self.currency)
        if self.currency==money.currency:
            return Currency(self.amount*money.amount, self.currency)
        else:
            logging.error("Before multiplying, please convert to the same currency")
            sys.exit(1)

    def __truediv__(self, money):
        """Si las divisas son distintas, queda el resultado con la divisa del primero"""
        if self.currency==money.currency:
            return Currency(self.amount/money.amount, self.currency)
        else:
            logging.error("Before true dividing, please convert to the same currency")
            sys.exit(1)

    def __repr__(self):
        return self.string(2)

    ## Returs a typical currency string
    ## @param digits int that defines the number of decimals. 2 by default
    ## @return string
    def string(self,   digits=2):
        return "{} {}".format(round(self.amount, digits), self.symbol())

    ## Returns the symbol of the currency
    def symbol(self):
        if self.currency=="EUR":
            return "€"
        elif self.currency=="USD":
            return "$"

    def isZero(self):
        if self.amount==Decimal(0):
            return True
        else:
            return False

    def isGETZero(self):
        if self.amount>=Decimal(0):
            return True
        else:
            return False

    def isGTZero(self):
        if self.amount>Decimal(0):
            return True
        else:
            return False

    def isLTZero(self):
        if self.amount<Decimal(0):
            return True
        else:
            return False

    def isLETZero(self):
        if self.amount<=Decimal(0):
            return True
        else:
            return False

    def __neg__(self):
        """Devuelve otro money con el amount con signo cambiado"""
        return Currency(-self.amount, self.currency)

    def round(self, digits=2):
        return round(self.amount, digits)

## Class to manage percentages in spreadsheets
class Percentage:
    def __init__(self, numerator=None, denominator=None):
        self.value=None
        self.setValue(self.toDecimal(numerator),self.toDecimal(denominator))

    def toDecimal(self, o):
        if o==None:
            return o
        if o.__class__==Currency:
            return o.amount
        elif o.__class__==Decimal:
            return o
        elif o.__class__ in ( int, float):
            return Decimal(o)
        elif o.__class__==Percentage:
            return o.value
        else:
            logging.warning (o.__class__)
            return None

    def __repr__(self):
        return self.string()

    def __neg__(self):
        if self.value==None:
            return self
        return Percentage(-self.value, 1)

    def __lt__(self, other):
        if self.value==None:
            value1=Decimal('-Infinity')
        else:
            value1=self.value
        if other.value==None:
            value2=Decimal('-Infinity')
        else:
            value2=other.value
        if value1<value2:
            return True
        return False

    def __mul__(self, value):
        if self.value==None or value==None:
            r=None
        else:
            r=self.value*self.toDecimal(value)
        return Percentage(r, 1)

    def __truediv__(self, value):
        try:
            r=self.value/self.toDecimal(value)
        except:
            r=None
        return Percentage(r, 1)

    def setValue(self, numerator,  denominator):
        try:
            self.value=Decimal(numerator/denominator)
        except:
            self.value=None

    def value_100(self):
        if self.value==None:
            return None
        else:
            return self.value*Decimal(100)

    def string(self, rnd=2):
        if self.value==None:
            return "None %"
        return "{} %".format(round(self.value_100(), rnd))

    ## Returns if the percentage is valid. I mean it's value different of None
    def isValid(self):
        if self.value!=None:
            return True
        return False

    def isGETZero(self):
        if self.value>=0:
            return True
        return False

    def isGTZero(self):
        if self.value>0:
            return True
        return False

    def isLTZero(self):
        if self.value<0:
            return True
        return False



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

## Class that manage spreadsheet coordinates (letter + number)
class Coord:
    def __init__(self, strcoord):
        self.letter, self.number=self.__extract(strcoord)
    def __repr__(self):
        return "Coord <{}>".format(self.string())

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

    ## Returns Coord string
    ## @return string For example "Z1"
    def string(self):
        return self.letter+self.number

    ## Add a number of rows to the Coord and return itself
    ## @param num Integer Can be positive and negative. When num is negative, if Coord.number is less than 1, returns 1
    def addRow(self, num=1):
        if self.numberIndex()+num<0:
            self.number="1"
        else:
            self.number=rowAdd(self.number, num)
        return self

    ## Add a number of columns/letters to the Coord and return itself
    ## @param num Integer. Can be positive and negative. When num is negative, if Coord.letter is less than A, returns A.
    def addColumn(self, num=1):
        if self.letterIndex()+num<0:
            self.letter="A"
        else:
            self.letter=columnAdd(self.letter, num)
        return self

    ## Reterns the letter/column index (letterPosition()-1)
    ## @return int
    def letterIndex(self):
        return column2index(self.letter)

    ## Returns the letter/column position
    ## @return int 
    def letterPosition(self):
        return column2number(self.letter)

    ## Returns the number/row index (numberPosition()-1)
    def numberIndex(self):
        return row2index(self.number)

    ## Returns the number/row position
    def numberPosition(self):
        return row2number(self.number)

    @staticmethod
    def assertCoord(o):
        if o.__class__==Coord:
            return o
        elif o.__class__==str:
            return Coord(o)


## Class that manages spreadsheet Ranges for ods and xlsx
class Range:
    def __init__(self,strrange):
        self.start, self.end=self.__extract(strrange)

    ## Converts a string to a Range. Returns None if conversion can't be done
    def __extract(self,range):
        if range.find(":")==-1:
            print("I can't manage this range")
            return
        a=range.split(":")
        return (Coord(a[0]), Coord(a[1]))

    ## String of a range in spreadsheets
    def string(self):
        return "{}:{}".format(self.start.string(), self.end.string())

    ## Number of range of the range
    def numRows(self):
        return row2number(self.end.number)-row2number(self.start.number) +1

    ## Number of columns of the range
    def numColumns(self):
        return column2number(self.end.letter)-column2number(self.start.letter) +1

    ## Returns a Range object even o is a str or a Range
    @staticmethod
    def assertRange(o):
        if o.__class__==Range:
            return o
        elif o.__class__==str:
            return Range(o)


    ## Adds a row to the end Coord, so it adds a row to the range
    def addRowAfter(self, num=1):
        self.end=self.end.addRow(num)
        return self

    ## Adds a column to the end Coord, so it adds a column to the range
    def addColumnAfter(self, num=1):
        self.end=self.end.addColumn(num)
        return self

    ## Adds a row to the top of the start Coord, so it adds a row to the range. If start Coord number is 1 returns the same Coord
    def addRowBefore(self, num=1):
        self.start=self.start.addRow(-num)
        return self

    ## Adds a column to the end Coord, so it adds a column to the range. If start Coord letter is A, returns the same Coord
    def addColumnBefore(self, num=1):
        self.start=self.start.addColumn(-num)
        return self
