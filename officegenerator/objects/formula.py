from .percentage import Percentage
from .currency import Currency
from decimal import Decimal
class Formula:
    ## Este objeto puede aañdir el valor con string_value and string_type o con value
    ## string_formula debue estar 
    def __init__(self, string_formula=None):
        self.string_formula=string_formula
        
    def __repr__(self):
        return "Formula ({})".format(self.string_formula)
        
    def hasObject(self):
        return hasattr(self, "object")
        
    ## Si no existe _object es porque no se ha establecido el valor. Si existe y es NOne es que es None.
    def setObject(self, object):
        self.object=object
        
    def setObjectFromStrings(self, string_value, string_type):
        self.object=string_value
        

## Returns true if value is a string beginning with = or +
## @param value must be a string
## @return boolean
def isFormula(value):
    if value.__class__.__name__=="str" and len(value)>0 and value[0] in ["=", "+"]:
        return True
    return False

## Generate a formula from a OdfpyCell
def Formula_from_OdfpyCell(odfpycell, string_type=None):
    r=Formula(odfpycell.getAttribute('formula'))
    textvalue=odfpycell.getAttribute('value')
    valuetype=odfpycell.getAttribute('valuetype')
    if odfpycell.getAttribute('valuetype')=="float":
        r.setObject(float(textvalue))
    elif odfpycell.getAttribute('valuetype')=="percentage":
        r.setObject(Percentage(Decimal(textvalue), 1))
    elif odfpycell.getAttribute('valuetype')=="currency":
        print(odfpycell.getAttribute("currency"))
        r.setObject(Currency(Decimal(textvalue), odfpycell.getAttribute("currency")))
    else:
        r.setObject(textvalue)
    print("Creando formula", textvalue, valuetype, "devolviendo", r.object, r.object.__class__)
    return r
