from logging import warning
from officegenerator.commons import Coord, index2column, index2row, columnAdd
from officegenerator.casts import lor_remove_columns, lor_remove_rows,  list_remove_positions, lor_add_column, lor_get_column, lor_get_row
from officegenerator.libodfgenerator import guess_ods_style, ODS_Write
from officegenerator.libxlsxgenerator import XLSX_Write
from officegenerator.objects.currency import currency_symbol

## Models are used to generate standard sheets
## We will use them to generate fast sheets with or without totals
## See the __main__ for examples

class ModelStyles:
    hh=0# Only horizontal header
    hv=1
    hht=2 #Only horizontal header with last row of totals
    hhv=3 # Only vertical header with last column of totals
    hhthhv=4 # Horizontal and vertical headers with last row and column of totals

class Model:
    def __init__(self):
        self.hh=None
        self.vh=None
        self.ht_definition=None
        self.vt_definition=None

    def setTitle(self, title):
        self.title=title
        
    ## If you want only vertical headers you must add self.setHorizontalHeaders(None, sizes=10)
    ## @param hh List or NOne
    ## @param sizes int or List in cm, to generate document use, columnSizes_for_ods
    def setHorizontalHeaders(self, hh, sizes=10):
        self.hh=hh
        if hh==None:#Only vertical headers
            columns=100
        else:
            columns=len(self.hh)+1
        if sizes.__class__.__name__=="int":
            self.hh_sizes=[sizes]*columns
        else:
            self.hh_sizes=sizes

    ## @param hh List or NOne
    ## @param size int with the witdth in cm of the vertical header
    def setVerticalHeaders(self, vh, size=5):
        self.vh=vh
        self.vh_size=size

#    ## REturns if a row is operable. It's used to generate automatic operable models
#    def isRowOperable(self, row_index, operable_from):
#        r=True
#        for cell in self.data[row_index]:
#            print( operable_from, row_index, cell,   cell.__class__.__name__)
#            if cell.__class__.__name__ in ["str",  "datetime", "date"]:
#                r=False
#                break
#        return r
#            
#    ## REturns if a row is operable. It's used to generate automatic operable models
#    def isColumnOperable(self, column_index, operable_from):
#        r=True
#        for i, row in enumerate(self.data):
#            if i>=operable_from and row[column_index].__class__.__name__ in ["str",  "datetime", "date" ]:
#                r= False
#                break
#        return r
        

    ## Order data columns. None values are set at the beginning
    def order_with_none(self, column_index, reverse=False, none_at_top=True):
        nonull=[]
        null=[]
        for row in self.data:
            if row[column_index] is None:
                null.append(row)
            else:
                nonull.append(row)
        nonull=sorted(nonull, key=lambda c: c[column_index], reverse=reverse)
        if none_at_top==True:#Set None at top of the list
            self.data= null+nonull
        else:
            self.data=nonull+null

    ## Converts self.hh_sizes and self.hv_size in cm to ods sizes. It returns them together
    def columnSizes_for_ods(self):
        r=[]
        factor=30
        if self.vh is not None:
            r.append(self.vh_size*factor)
        for arg in self.hh_sizes:
                r.append(arg*factor)
        return r
        
    ## Returns the number of rows in data
    def numDataRows(self):
        return len(self.data)

    ## Returns the number of rows in data
    def numDataColumns(self):
        return len(self.data[0])
        
    ## Converts self.hh_sizes in cm to xlsx sizes
    def columnSizes_for_xlsx(self):
        r=[]
        factor=6
        if self.vh is not None:
            r.append(self.vh_size*factor)
        for arg in self.hh_sizes:
                r.append(arg*factor)
        return r
        
    ## Converts self.hh_sizes in cm to odt sizes,proporcional to the parameter
    def columnSizes_for_odt(self, tablesize):
        cm=[]
        factor=1
        if self.vh is not None:
            cm.append(self.vh_size*factor)
        for arg in self.hh_sizes:
                cm.append(arg*factor)
        #Until here are cm but can oversize the maximum
        r=[]
        sum_=sum(cm)
        for size in cm:
            r.append(tablesize*size/sum_)
        return r

    def setData(self, data):
        self.data=data
        
    ## Used to remove Columns in the self.data
    ## @param columnList List of integers with the column index to remove
    def removeColumns(self, columnList):
        if self.hh is not None:
            self.hh=list_remove_positions(self.hh, columnList)
        self.hh_sizes=list_remove_positions(self.hh_sizes, columnList)
        if self.data is not None:
            self.data=lor_remove_columns(self.data, columnList)
        else:
            warning("I can't remove columns if self.data is None")

    ## Used to remove Columns in the self.data
    ## @param columnList List of integers with the column index to remove
    def removeRows(self, rowList):
        if self.vh is not None:
            self.vh=list_remove_positions(self.vh, rowList)
        if self.data is not None:
            self.data=lor_remove_rows(self.data, rowList)
        else:
            warning("I can't remove rows if self.data is None")

    ## @param title String with the title of the sheet
    ## @param columns_title List of Strings
    ## @param data list of list
    ## @param sizes List of integers
    def xlsx_sheet(self, doc):
        doc.createSheet(self.title)
        doc.setColumnsWidth(self.columnSizes_for_xlsx())
        if self.hh is not None:
            doc.overwrite(self.__getFirstContentCoord().addRow(-1), [self.hh], doc.stOrange)
        if self.vh is not None:
            for i, header in enumerate(self.vh):
                doc.overwrite(self.__getFirstContentCoord().addColumn(-1).addRow(i), header, doc.stGreen, alignment="left")
        if self.__mustFillA1()==True:
            doc.overwrite("A1", " ",  doc.stOrange)

        if self.numDataRows()>0:
            for number, row in enumerate(self.data):
                for letter,  field in enumerate(row):
                    doc.overwrite(self.__getFirstContentCoord().addRow(number).addColumn(letter), field, style=doc.stWhite)

            #Fills horizontal  total
            if self.ht_definition is not None:
                for letter, definition in enumerate(self.ht_definition):
                    class_=self.__object_to_formula_classname(self.data[0][letter])
                    if self.__is_total_key(definition):
                        doc.overwrite_formula(self.__getFirstContentCoord().addRow(self.numDataRows()).addColumn(letter), self.__calculate_horizontal_total("xlsx", letter), class_,  style=doc.stGrayLight)
                    else:
                        doc.overwrite(self.__getFirstContentCoord().addRow(self.numDataRows()).addColumn(letter), self.__calculate_horizontal_total("xlsx", letter), style=doc.stGrayLight)
            #Fills vertical total
            if self.vt_definition is not None:
                for number, definition in enumerate(self.vt_definition):
                    class_=self.__object_to_formula_classname(self.data[0][self.numDataColumns()-1])
                    value=self.__calculate_vertical_total("xlsx", number)
                    if self.__is_total_key(definition):
                        doc.overwrite_formula(Coord("A1").addRow(number).addColumn(self.numDataColumns()), value, class_,  style=doc.stGrayLight)
                    else:
                        doc.overwrite(Coord("A1").addRow(number).addColumn(self.numDataColumns()), value, style=doc.stGrayLight)
        self.__setFreezeAndSelect(doc)

    ## Function neeeded to change formula types, due to is a string but needs to be changed to currency symbol. to use __setFormulaNumberFormat
    def __object_to_formula_classname(self, o):
        if o.__class__.__name__ in ("Currency", "Money"):
            return currency_symbol(o.currency) 
        else:
            return o.__class__.__name__

    ## @param title String with the title of the sheet
    ## @param columns_title List of Strings
    ## @param data list of list
    ## @param sizes List of integers
    def ods_sheet(self, doc):
        s=doc.createSheet(self.title)
        s.setColumnsWidth(self.columnSizes_for_ods())
        if self.hh is not None:
            s.add(self.__getFirstContentCoord().addRow(-1), [self.hh],  "OrangeCenter")
        if self.vh is not None :
            for i, header in enumerate(self.vh):
                s.add(self.__getFirstContentCoord().addColumn(-1).addRow(i), header, "GreenLeft")
        if self.__mustFillA1()==True:
            s.add("A1", "", "OrangeCenter")

        if self.numDataRows()>0: #Only must be executed with data
            for number, row in enumerate(self.data):
                for letter,  field in enumerate(row):
                    s.add(self.__getFirstContentCoord().addRow(number).addColumn(letter), field)
                    
            #Fills horizontal 
            if self.ht_definition is not None:
                for letter, definition in enumerate(self.ht_definition):
                    value= self.__calculate_horizontal_total("ods", letter)
                    style=guess_ods_style("GrayLight", self.data[self.ht_index_from][letter])
                    s.add(Coord("A1").addRow(self.numDataRows()+1).addColumn(letter), value, style )
            #Fills vertical
            if self.vt_definition is not None:
                for number, definition in enumerate(self.vt_definition):
                    value=self.__calculate_vertical_total("ods", number)
                    style=guess_ods_style("GrayLight", self.data[self.vt_index_from][letter])
                    s.add(Coord("A1").addRow(number).addColumn(self.numDataColumns()), value,  style)

        self.__setFreezeAndSelect(s)
        return s

    ## @param type can be "ods","xlsx","odt","value"
    ## See setHorizontalTotalDefinition doc for available keys
    def __calculate_horizontal_total(self, type, column_index):
        key=self.ht_definition[column_index]
        column=index2column(column_index)
        total_from=index2row(self.ht_index_from)
        r=key
        if type in("ods",  "xlsx"):
            if key=="#SUM":
                r= "=SUM({0}{1}:{0}{2})".format(column, total_from, self.numDataRows()+1)
            elif key=="#AVG":
                r= "=AVERAGE({0}{1}:{0}{2})".format(column, total_from, self.numDataRows()+1)
            elif key=="#MEDIAN":
                r= "=MEDIAN({0}{1}:{0}{2})".format(column, total_from, self.numDataRows()+1)
        elif type=="value":
            if key=="#SUM":
                r= sum(lor_get_column(self.data, column))
        return r

    ## @param type can be "ods","xlsx","odt","value"
    ## See setHorizontalTotalDefinition doc for available keys
    def __calculate_vertical_total(self, type, row_index):
        key=self.vt_definition[row_index]
        row=index2row(row_index)
        total_from=index2column(self.vt_index_from)
        r=key
        if type in("ods",  "xlsx"):
            if key=="#SUM":
                r= "=SUM({0}{1}:{2}{1})".format(total_from, row, columnAdd(total_from, self.numDataColumns()-self.vt_index_from-1))
            elif key=="#AVG":
                r= "=AVERAGE({0}{1}:{2}{1})".format(total_from, row, columnAdd(total_from, self.numDataColumns()-self.vt_index_from-1))
            elif key=="#MEDIAN":
                r= "=MEDIAN({0}{1}:{2}{1})".format(total_from, row, columnAdd(total_from, self.numDataColumns()-self.vt_index_from-1))
        elif type=="value":
            if key=="#SUM":
                r= sum(lor_get_row(self.data, row))
        return r
        
        
    def __is_total_key(self, s):
        if s in ["#SUM","#AVG","#MEDIAN"]:
            return True
        return False

    ## Generates a odt table object from model
    ## @param doc odt document
    ## @param tablesize float in cm where the table is going to be placed in the paper(document)
    ## @param fontsiz1e int withe the size of the font in the document
    ## @param after officegenerator after parameter
    def odt_table(self, doc, tablesize, fontsize, after=True):
        if self.vh is not None:
            data=lor_add_column(self.data, 0, self.vh)
        else:
            data=self.data
        if self.__mustFillA1()==True:
            hh=[" ", ] + self.hh
        else:
            hh=self.hh
        return doc.table(hh, data, self.columnSizes_for_odt(tablesize), fontsize, self.title, after)

    def __getFirstContentCoord(self):
        #firstcontentletter and firstcontentnumber
        if self.hh is None and self.vh is not None:
            return Coord("B1")
        elif self.hh is not None and self.vh is None:
            return Coord("A2")
        elif self.hh is not None and self.vh is not None:
            return Coord("B2")
        elif self.hh is None and self.vh is None:
            return Coord("A1")
        
    ## @param ref Reference to object: ods: sheet, xlsx: doc
    def __setFreezeAndSelect(self, ref):
        if self.ht_definition is not None:
            ref.freezeAndSelect(self.__getFirstContentCoord(), self.__getFirstContentCoord().addRow(self.numDataRows()).addColumn(self.numDataColumns()-1))
        elif self.vt_definition is not None:
            ref.freezeAndSelect(self.__getFirstContentCoord(), self.__getFirstContentCoord().addRow(self.numDataRows()-1).addColumn(self.numDataColumns()))
        else:
            ref.freezeAndSelect(self.__getFirstContentCoord(),self.__getFirstContentCoord().addRow(self.numDataRows()-1).addColumn(self.numDataColumns()-1))
            
        
    ## Return if A1 must be filled for a better view
    ## @param bool
    def __mustFillA1(self):
        if self.hh is not None and self.vh is not None:
            return True
        return False
        
    
    ## Available keys:
    ## - #SUM
    ## - #AVG
    ## - #COUNT
    ## - #MEDIAN
    ## @param definition_list List with strings and keys
    ## @param totals_index_from Column index from with totals are generated
    def setHorizontalTotalDefinition(self, definition_list, totals_index_from=1):
        self.ht_definition=definition_list    ## Available keys:
        self.ht_index_from=totals_index_from

    ## See setHorizontalTotalDefinition doc for available keys
    ## @param definition_list List with strings and keys
    ## @param totals_index_from Column index from with totals are generated
    def setVerticalTotalDefinition(self, definition_list, totals_index_from=1):
        self.vt_definition=definition_list
        self.vt_index_from=totals_index_from
        
    ## Generates an xlsx file from model
    def xlsx_file(self, filename):
        xlsx=XLSX_Write(filename)
        self.xlsx_sheet(xlsx)
        xlsx.remove_sheet_by_id(0)
        xlsx.save()

    ## Generates an ods file from model
    def ods_file(self, filename):
        ods=ODS_Write(filename)
        self.ods_sheet(ods)
        ods.save()
        
        
## Can show ht and vt. from an index
## @param title Title of the model and sheet
## @param hh List with Horizontal headers
## @param data lor with model data
## @param allkeys Operation key for all fields
## @param row_index_from Defines if some rows are not going to be operated in formulas
## @param column_index_from Defines if some columns are not going to be operated in formulas
## @param ht Boolean to show horizontal totals
## @param vt Boolean to show vertical totals
## @param skip_ht List of index column to skip some totals
## @param skip_vt List of index rows to skip some totals
def Model_Auto(title, hh, data, allkeys="#SUM", sizes=5, row_index_from=1, column_index_from=1, ht=True, vt=True, skip_ht=[],skip_vt=[]):
    m=Model()
    m.setTitle(title)
    m.setHorizontalHeaders(hh, sizes)
    m.setData(data)

    if ht is True:
        hdef=["Total", ]
        hdef=hdef+[""]*(column_index_from-1)# Empty totals
        for index in range(column_index_from, m.numDataColumns()):
            if index in skip_ht:
                hdef=hdef+[""]
            else:
                hdef=hdef+[allkeys]
        m.setHorizontalTotalDefinition(hdef, totals_index_from=column_index_from)

    if vt is True:
        vdef=["Total", ]
        vdef=vdef+[""]*(row_index_from-1)# Empty totals
        for index in range(row_index_from, m.numDataRows()+1):
            if index in skip_vt:
                vdef=vdef+[""]
            else:
                vdef=vdef+[allkeys]
        
        if ht is True:
            vdef=vdef + [allkeys]#Crossed total
            
        m.setVerticalTotalDefinition(vdef, totals_index_from=column_index_from)
    return m
