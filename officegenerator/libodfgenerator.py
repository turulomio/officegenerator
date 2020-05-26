## @namespace officegenerator.libodfgenerator
## @brief Package that allows to read and write Libreoffice ods and odt files
## This file is from the Xulpymoney project. If you want to change it. Ask for project administrator

## odf Element and Text inherits from Node and inherits from xml.Dom.Node https://docs.python.org/3.5/library/xml.dom.html
## Text can't have children

from datetime import datetime, time
import gettext
from decimal import Decimal
from logging import info, debug
from odf.opendocument import OpenDocumentSpreadsheet,  OpenDocumentText,  load
from odf.style import Footer, FooterStyle, HeaderFooterProperties, Style, TextProperties, TableColumnProperties, Map,  TableProperties,  TableCellProperties, PageLayout, PageLayoutProperties, ParagraphProperties,  ListLevelProperties,  MasterPage
from odf.number import  TimeStyle, CurrencyStyle, CurrencySymbol,  Number, NumberStyle, Text,  PercentageStyle,  DateStyle, Year, Month, Day, Hours, Minutes, Seconds, Boolean, BooleanStyle
from odf.text import P,  H,  Span, ListStyle,  ListLevelStyleBullet,  List,  ListItem, ListLevelStyleNumber,  OutlineLevelStyle,  OutlineStyle,  PageNumber,  PageCount
from odf.table import Table, TableColumn, TableRow, TableCell,  TableHeaderRows
from odf.draw import Frame, Image
from odf.dc import Creator, Description, Title, Date, Subject
from odf.meta import InitialCreator, Keyword, CreationDate
import odf.element

from odf.config import ConfigItem, ConfigItemMapEntry, ConfigItemMapIndexed, ConfigItemMapNamed,  ConfigItemSet
from odf.office import Annotation
from officegenerator.commons import number2column,  number2row,  Coord, Range, topLeftCellNone, column2index, Coord_from_index, row2index, convert_command, generate_formula_total_string
from officegenerator.objects.currency import Currency
from officegenerator.objects.formula import Formula_from_OdfpyCell, isFormula, Formula
from officegenerator.datetime_functions import dtnaive2string
from officegenerator.objects.percentage import Percentage
from os import path, remove, makedirs, sep
from pkg_resources import resource_filename
from tempfile import TemporaryDirectory

try:
    t=gettext.translation('officegenerator', resource_filename("officegenerator","locale"))
    _=t.gettext
except:
    _=str

## Type class that defines predefined width columns
class ColumnWidthODS:
    Date=60
    Datetime=100
    XS=12
    S=25
    M=50
    L=100
    XL=150
    XXL=200
    XXXL=250
    XXXXL=300

class ODS_Read:
    def __init__(self, filename):
        self.doc=load(filename)#doc is only used in this function. All is generated in self.doc
        self.filename=filename

    def getSheetElementByName(self, name):
        """
            Devuelve el elemento de sheet buscando por su nombre
        """
        for numrow, sheet in  enumerate(self.doc.spreadsheet.getElementsByType(Table)):
            if sheet.getAttribute("name")==name:
                return sheet
        return None        

    def getSheetElementByIndex(self, index):
        """
            Devuelve el elemento de sheet buscando por su posición en el documento
        """
        try:
            return self.doc.spreadsheet.getElementsByType(Table)[index]
        except:
            return None
#        
#    ## Returns a list of rows with the odfcells of the sheet
#    ## @param sheet_index Integer index of the sheet
#    ## @param skip_up int. Number of rows to skip at the begining of the list of rows (lor)
#    ## @param skip_down int. Number of rows to skip at the end of the list of rows (lor)
#    ## @return Returns a list of rows of object values
#    def cells(self, sheet_index, skip_up=0, skip_down=0):
#        sheet_element=self.getSheetElementByIndex(sheet_index)        
#        rows=sheet_element.getElementsByType(TableRow) #Uses ODFPY cell to boost performance
#        r=[]
#        for row in rows[skip_up:len(rows)-skip_down]:
#            tmprow=[]
#            for cell in row.getElementsByType(TableCell):
#                tmprow.append(self._getCellValue_from_odfpy_cell(cell))
#            r.append(tmprow)
#        return r

        
#    ## REturns a lor of odfpy cells
#    def cellsOdfPy(self, sheet_index):
#        r=[]
#        sheet_element=self.getSheetElementByIndex(sheet_index)
#        for row in sheet_element.getElementsByType(TableRow):
#            r_row=[]
#            for cell in row.getElementsByType(TableCell):
#                r_row.append(cell)
#            r.append(r_row)
#        return r

    ## Returns a list of rows with the values of the sheet
    ## @param sheet_index Integer index of the sheet
    ## @param skip_up int. Number of rows to skip at the begining of the list of rows (lor)
    ## @param skip_down int. Number of rows to skip at the end of the list of rows (lor)
    ## @return Returns a list of rows of object values
    def values(self, sheet_index, skip_up=0, skip_down=0):
        sheet_element=self.getSheetElementByIndex(sheet_index)        
        rows=sheet_element.getElementsByType(TableRow) #Uses ODFPY cell to boost performance
        r=[]
        for row in rows[skip_up:len(rows)-skip_down]:
            tmprow=[]
            for cell in row.getElementsByType(TableCell):
                tmprow.append(self._getCellValue_from_odfpy_cell(cell))
            r.append(tmprow)
        return r

    ## @param sheet_index Integer index of the sheet
    ## @param range_ Range object to get values. If None returns all values from sheet
    ## @return Returns a list of rows of object values
    def values_by_range(self, sheet_index, range_):
        sheet_element=self.getSheetElementByIndex(sheet_index)        
        rows=sheet_element.getElementsByType(TableRow) #Uses ODFPY cell to boost performance
        r=[]
        range_=Range.assertRange(range_)
        for number_index, row in enumerate(rows):
            tmprow=[]
            for letter_index, cell in enumerate(row.getElementsByType(TableCell)):
                if Coord_from_index(letter_index, number_index) in range_:
                    tmprow.append(self._getCellValue_from_odfpy_cell(cell))
            r.append(tmprow)
        return r
    
    ## @param sheet_index Integer index of the sheet
    ## @param column_letter Letter of the column to get values
    ## @param skip Integer Number of top rows to skip in the result
    ## @return List of values
    def getColumnValues(self, sheet_index, column_letter, skip_up=0, skip_down=0):
        r=[]
        sheet_element=self.getSheetElementByIndex(sheet_index)        
        rows=sheet_element.getElementsByType(TableRow) #Uses ODFPY cell to boost performance
        for row in rows[skip_up:len(rows)-skip_down]:
            cell=row.getElementsByType(TableCell)[column2index(column_letter)]
            r.append(self._getCellValue_from_odfpy_cell(cell))
        return r    

    ## @param sheet_index Integer index of the sheet
    ## @param row_number String Number of the row to get values
    ## @param skip Integer Number of top rows to skip in the result
    ## @return List of values
    def getRowValues(self, sheet_index, row_number, skip_left=0, skip_right=0):
        r=[]
        sheet_element=self.getSheetElementByIndex(sheet_index)        
        #Uses ODFPY cell to boost performance
        row=sheet_element.getElementsByType(TableRow)[row2index(row_number)]
        cells=row.getElementsByType(TableCell)
        for cell in cells[skip_left:len(cells)-skip_right]:
            r.append(self._getCellValue_from_odfpy_cell(cell))
        return r    

    ## Return a Range object with the limits of the index sheet
    def getSheetRange(self, sheet_index):
        endcoord=Coord("A1").addRow(self.rowNumber(sheet_index)-1).addColumn(self.columnNumber(sheet_index)-1)
        return Range("A1:"+endcoord.string())
        
    def rowNumber(self, sheet_index):
        """
            Devuelve el numero de filas de un determinado sheet_element
        """
        sheet_element=self.getSheetElementByIndex(sheet_index)
        return len(sheet_element.getElementsByType(TableRow))
        
    def columnNumber(self, sheet_index):
        """
            Devuelve el numero de filas de un determinado sheet_element
        """
        sheet_element=self.getSheetElementByIndex(sheet_index)
        return len(sheet_element.getElementsByType(TableColumn))
        
    def getOdfPyCell(self, sheet_index, coord):
        coord=Coord.assertCoord(coord)
        sheet_element=self.getSheetElementByIndex(sheet_index)
        row=sheet_element.getElementsByType(TableRow)[coord.numberIndex()]
        return row.getElementsByType(TableCell)[coord.letterIndex()]
        
    ## Returns the cell value
    def getCellValue(self, sheet_index, coord):
        cell=self.getOdfPyCell(sheet_index, coord)
        return self._getCellValue_from_odfpy_cell(cell)
        
    ## Used to improve performance avoiding searching cells
    def _getCellValue_from_odfpy_cell(self, cell):
        r=None
        if cell.getAttribute('formula') is not None:
            r=Formula_from_OdfpyCell(cell)
        elif cell.getAttribute('valuetype')=='string':
            r=str(cell)
        elif cell.getAttribute('valuetype')=='float':
            r=Decimal(cell.getAttribute('value'))
        elif cell.getAttribute('valuetype')=='percentage':
            r=Percentage(Decimal(cell.getAttribute('value')), Decimal(1))
        elif cell.getAttribute('valuetype')=='currency':
            r=Currency(Decimal(cell.getAttribute('value')), cell.getAttribute('currency'))
        elif cell.getAttribute('valuetype')=='date':
            datevalue=cell.getAttribute('datevalue')
            if len(datevalue)<=10:
                r=datetime.strptime(datevalue, "%Y-%m-%d").date()
            else:
                r=datetime.strptime(datevalue, "%Y-%m-%dT%H:%M:%S")
        elif cell.getAttribute('valuetype')=='time':
            s=cell.getAttribute('timevalue')
            h=int(s.split("PT")[1].split("H")[0])
            m=int(s.split("H")[1].split("M")[0])
            s=int(s.split("M")[1].split("S")[0])
            r=time(h, m, s)
        else:
            return None
        return r

    ## Returns an odfcell object
    def getCell(self, sheet_index,  coord):
        cell=self.getOdfPyCell(sheet_index, coord)
        object=self._getCellValue_from_odfpy_cell(cell)
        #Get spanning
        spanning_columns=cell.getAttribute('numbercolumnsspanned')
        if spanning_columns==None:
            spanning_columns=1
        else:
            spanning_columns=int(spanning_columns)
        spanning_rows=cell.getAttribute('numberrowsspanned')
        if spanning_rows==None:
            spanning_rows=1
        else:
            spanning_rows=int(spanning_rows)
        
        #Get Stylename
        stylename=cell.getAttribute('stylename')


        #Odfcell
        r=OdfCell(coord, object, stylename)
        r.setSpanning(spanning_columns, spanning_rows)
        return r
        
    def setCell(self, sheet_index,  coord, odfcell):
        """
            odfcell is a officegenerator.OdfCell object
            Updates a cell
            insertBefore(newchild, refchild) – Inserts the node newchild before the existing child node refchild.
appendChild(newchild) – Adds the node newchild to the end of the list of children.
removeChild(oldchild) – Re

        ESTA FUNCION SE USA PARA SUSTITUIR EN UNA PLANTILLA
        NO SE PUEDEN AÑADIR MAS CELDAS O FILAS
        PARA ESO USAR ODS_Write DE MOMENTO
        """
        coord=Coord.assertCoord(coord)
        sheet_element=self.getSheetElementByIndex(sheet_index)
        row=sheet_element.getElementsByType(TableRow)[coord.numberIndex()]
        oldcell=row.getElementsByType(TableCell)[coord.letterIndex()]
        row.insertBefore(odfcell.generate(), oldcell)
        row.removeChild(oldcell)        



    def save(self, filename):
        if  filename==self.filename:
            print(_("You can't overwrite a readed ods"))
            return        
        self.doc.save( filename)


## Abstract class the defines opendocument metadata, images, languages
class ODF:
    def __init__(self, filename):
        self.filename=filename
        self.images={}
        
    ## Set metadata in odf document
    ## @param title String with the title of the document
    ## @param subject String with a brief of the document
    ## @param creator String with the author of the document
    ## @param description String with a larga description of the document
    ## @param keywords String with keywords separated by space
    ## @param creationdate Naive datetime with the creation date and time
    def setMetadata(self, title="",  subject="", creator="", description="", keywords="", creationdate=datetime.now()):
        for e in self.doc.meta.childNodes:
            self.doc.meta.removeChild(e)
        self.doc.meta.addElement(Description(text=description))
        self.doc.meta.addElement(Title(text=title))
        self.doc.meta.addElement(Subject(text=subject))
        self.doc.meta.addElement(Creator(text=creator))
        self.doc.meta.addElement(InitialCreator(text=creator))
        self.doc.meta.addElement(Keyword(text=keywords))
        d=Date()
        d.addText(creationdate.strftime("%Y-%m-%dT%H:%M:%S"))
        self.doc.meta.addElement(CreationDate(text=d))

    ## Adds an image to self.images dictionary. We add to a dictionary in order to reuse the same image in a directory
    ##
    ## One in the directory you can use the image with ODT.image()
    ## @param path String with the path to the image
    ## @param key String with the key of self.images dictionary. By default is None, in this case, the path string will be the key.
    def addImage(self, path, key=None):
        if key==None:
            key=path
        self.images[key]=self.doc.addPicture(path)

    def setLanguage(self, language, country):
        """Set the main language of the document"""
        self.language="es"
        self.country="ES"
        
    def showElement(self, e):
        print("ATTRIBUTE_NODE: {}".format(e.ATTRIBUTE_NODE))
        print("CDATA_SECTION_NODE: {}".format(e.CDATA_SECTION_NODE))
        print("TEXT_NODE: {}".format(e.TEXT_NODE))
        print("Atributes: {}".format(e.attributes))
        print("QNAME: {}".format(e.qname))
        print("tagName: {}".format(e.tagName))
        print("Allowed attributes: {}".format(e.allowed_attributes()))
        
        print(e)

## Class used to generate a ODT file with predefined formats
## @param filename String with the name of the filename to read, then will be saved with a different name
## @param template String with the name of the filename used as template
## @param language String with language. For example es
## @param country String with the country. For example ES
## @param predefinedstyles Boolean that sets if predefined styles are going to be loaded. True by default
class ODT(ODF):
    def __init__(self, filename, template=None, language="es", country="ES"):
        ODF.__init__(self, filename)
        self.setLanguage(language, country)
        ## After inserting an element it sets the new element as cursor to append 

        self.seqTables=0#Sequence of tables
        self.seqFrames=0#If a frame is repeated it doesn't show its
        self.template=template
        if self.template==None:
            self.doc=OpenDocumentText()
        else:
            self.doc= load(self.template)
            self.save()#After loading a template it's needed to save file to work with search functions
            remove(self.filename)#Delete temporal save
        self.cursor=None
        self.cursorParent=self.doc.text

    ## Creates a text header
    ## @param text String with the header string
    ## @Level Integer Level of the header
    def header(self, text, level, after=True):
        h=H(outlinelevel=level, stylename="Heading_20_{}".format(level), text=text)
        return self.insertInCursor(h, after)

    def pageBreak(self,  horizontal=False, after=True):    
        p=P(stylename="PageBreak")#Is an automatic style
        self.doc.text.addElement(p)
        if horizontal==True:
            p=P(stylename="PH")
        else:
            p=P(stylename="PV")
        return self.insertInCursor(p, after)


    ## @param href must bu added before with addImage
    ## @param width Int or float value
    ## @param height Int or float value to set the images height
    ## @return Frame element
    def image(self, href, width, height, name=None):
        self.seqFrames=self.seqFrames+1
        name=name if name!=None else "Frame.{}".format(self.seqFrames)

        f = Frame(name=name, anchortype="as-char", width="{}cm".format(width), height="{}cm".format(height))
        img = Image(href=self.images[href], type="simple", show="embed", actuate="onLoad")
        f.addElement(img)
        return f

    ## List_20_1 is a list style. Don't get wrong with List_20_1 paragraph style
    ## @param arr List of strings and other list. You must take care that level2 list are included in list-item of level 1.
    ## @param style is a list style not a paragraph style. In standard configured styles are "List_20_1" and "Numbering_20_123"
    ## @param boolean To insert before or after the current cursor.
    ## @code multilevel
    ##    doc.list(   [   ["1", ["1.1", "1.2"]], 
    ##                    ["2"], 
    ##                    ["3",  ["3.1", ["3.1.1", "3.1.2"]]]
    ##                ],  style="List_20_1")   
    ## @endcode
    ## @code unilevel
    ##    doc.list(   [   ["1",], 
    ##                        ["2",], 
    ##                        ["3",],
    ##                ],  style="List_20_1")   
    ## @endcode
    ## Se genera un nivel más de list no se porque pero queda bien.
    def list(self, arr, list_style="List_20_1", paragraph_style="Text_20_body", after=True):
        def get_items(list_o, list_style, paragraph_style):
            r=[]
            for o in list_o:
                it=ListItem()
                if o.__class__==str:
                    it.addElement(P(stylename=paragraph_style, text=o))
                else:
                    it.addElement(get_list(o, list_style, paragraph_style))
                r.append(it)
            return r
        def get_list(arr, list_style, paragraph_style):
            ls=List(stylename=list_style)
            for listitem in get_items(arr, list_style, paragraph_style):
                ls.addElement(listitem)
            return ls
        # #########################
        return self.insertInCursor(get_list(arr, list_style, paragraph_style), after)    

    ## Extracts odf document structure
    def odf_dump_nodes(self, start_node, level=0):
        if start_node.nodeType==3:
            # text node
            print ("  "*level, "NODE:", start_node.nodeType, ":(text):")
        else:
            # element node
            attrs= []
            for k in start_node.attributes.keys():
                attrs.append( str(k[1]) + ':' + str(start_node.attributes[k]  ))
            print("{} NODE: {}:{} ATTR: {}".format(" "*level,  str(start_node.nodeType), str(start_node.qname[1]), attrs))

            for n in start_node.childNodes:
                self.odf_dump_nodes(n, level+1)

    ## Adds a paragraph
    ## @param text to add
    ## @param style Style to use
    ## @param after True: insert after self.cursor element. False: insert before self.cursor element. None: Just return element
    def simpleParagraph(self, text, style="Standard", after=True):
        p= P(stylename=style, text=text)
        return self.insertInCursor(p, after)

    ## Inserts after or before the Cursor, and sets the Cursor to the o element
    ## @param o Object to insert
    ## @param after True: insert after self.cursor element. False: insert before self.cursor element. None: Just return element
    def insertInCursor(self, o, after):
        if after==None:# It doesn't insert, just return item
            return o

        if self.cursor==None:# First insert
            self.cursorParent.addElement(o)
        elif after==True:
            indexcursor=self.cursorParent.childNodes.index(self.cursor)
            if len(self.cursorParent.childNodes)==indexcursor+1:
                self.cursorParent.insertBefore(o, None)
            else:
                siguiente=self.cursorParent.childNodes[indexcursor+1]
                self.cursorParent.insertBefore(o, siguiente)
        else:#After = False
            self.cursorParent.insertBefore(o, self.cursor)
        self.__setCursor(o)
        return o

    ## Search for a tag type_ elementes  an returns element and its index. With it, inserts new with replaced text and removes old
    ##
    ## Cursor doesn't change because we replace Text objects in Element Text
    ## @param tag String to search
    ## @param replace String to replace. Can't be None
    ## @param Search in elements with tag P or H. To encapsulate type_ will be a string "P" or "H"
    def search_and_replace(self, tag, replace, type_="P"):
        e,  textindex=self.search(tag, type_) #Places cursor to element
        if e==None:
            return

        if replace==None:#Remove paragraph
            print("Replace parameter can't be None. Use '' instead or use search_and_replace_element")
            return

        #Replace text with the same style as found paragraph
        to_remove=e.childNodes[textindex]
        if replace==None:#Removes text
            e.removeChild(to_remove)
            if len(e.childNodes)==0:#Removes element
                self.cursorParent.removeChild(e)
        else: #Replace
            e.insertBefore( odf.element.Text(str(to_remove).replace(tag, replace)), to_remove)
            #print(to_remove.__class__, e.__class__, e.childNodes)
            debug("THIS CODE FAILS IN ODFPY-1.4.1 AND WORKS IN ODFPY-1.3.6")
            e.removeChild(to_remove     )
            

    ## Search for a tag in doc an replaces its elemente with the parameter element
    ##
    ## @param tag String to search
    ## @param replace ELement. OdfPy element
    def search_and_replace_element(self, tag, newelement, type_="P"):
        e,  textindex=self.search(tag, type_) #Places cursor to element
        if e==None:
            return

        if newelement==None:#Remove paragraph
            print("New element can't be None")
            return

        self.cursorParent.insertBefore(newelement,e)
        self.cursorParent.removeChild(e)
        self.__setCursor(newelement)


    ## Searchs for the item with a tag. Perhaps is its paren where I'll have to append. Only finds the first one
    ## Returns the element p and the position in its text children
    ## Using templates sometimes you cant search a tag. It's due to sometimes has <span> in the tag. Use odf2xml to detect them
    ## 20200119 Search_and_replace has problem with tables, we must try to not abuse of them
    def search(self, tag, type_="P"):
        if type_=="P":
             ty=P
        elif type_=="H":
             ty=H
        for e in self.doc.getElementsByType(ty):
            #print("Searching", tag, "found", str(e))
            if str(e).find(tag)!=-1:
                self.__setCursor(e)
                for index, child in enumerate(e.childNodes):
                    #print(index, child)
                    #print("Searching in for", tag,"-", child , "found", str(e))
                    if str(child).find(tag)!=-1:
                        #print("SEARCH RETURN",  e, index)
                        return e, index
        print ("Tag {} not found with type {}".format(tag,type_))
        return None, None


    ## Converts saved odt to pdf. It will have the same basename but with .pdf extension
    ## @param output_dir None or dirname. If set, it writes pdf in that directory with the same basename of the filename class attribute
    def convert_to_pdf(self, output_dir=None):
        if output_dir is None:
            s_outdir=""
        else:
            s_outdir="--outdir '{}'".format(output_dir)
        convert_command(self.filename, s_outdir, "pdf")

    def __setCursor(self, e):
        self.cursor=e
        self.cursorParent=e.parentNode
        if self.cursorParent==None:
            print("Parent of '{}' is None".format(self.cursor))

    ## Check if self.filename is different to self.template, create directory and saves the file
    def save(self):
        if  self.filename==self.template:
            print(_("You can't overwrite a readed odt"))
            return
        if path.dirname(self.filename)!="":
            makedirs(path.dirname(self.filename), exist_ok=True)
        self.doc.save( self.filename)

    ## Adds an empty paragraph
    ## @param style String with the style name
    ## @param number Integer with the number of times to repeat the empty paragraph
    ## @param after True: insert after self.cursor element. False: insert before self.cursor element. None: Just return element
    def emptyParagraph(self, style="Standard", number=1, after=True):
        for i in range(number):
            self.simpleParagraph("",style, after)

    ## Creates the document title
    ## @param text String with the title
    ## @param after True: insert after self.cursor element. False: insert before self.cursor element. None: Just return element
    def title(self, text, after=True):
        p=P(stylename="Title", text=text)
        return self.insertInCursor(p, after)


    ## Creates the document title
    ## @param text String with the title
    ## @param after True: insert after self.cursor element. False: insert before self.cursor element. None: Just return element
    def subtitle(self, text, after=True):
        p=P(stylename="Subtitle", text=text)
        return self.insertInCursor(p, after)



    ## Creates a table adding it to self.doc
    ## @param hh List with all header strings
    ## @param data Multidimension List with all data objects. Can be str, Decimal, int, datetime, date, Currency, Percentage
    ## @param sizes Integer list with sizes in cm
    ## @param fontsize Integer in pt
    ## @param name str or None. Sets the object name. Appears in LibreOffice navigator. If none table will be named to "Table.Sequence"
    ## @param after True: insert after self.cursor element. False: insert before self.cursor element. None: Just return element
    def table(self, hh, data, sizes, fontsize, name=None, after=True):
        def generate_table_styles():
            s=Style(name="Table.Size{}".format(sum(sizes)), family='table')
            s.addElement(TableProperties(width="{}cm".format(sum(sizes)), align="center", margintop="0.6cm", marginbottom="0.6cm"))
            self.doc.automaticstyles.addElement(s)

            #Column sizes
            for i, size in enumerate(sizes):
                sc= Style(name="Table.Column.Size{}".format(size), family="table-column")
                sc.addElement(TableColumnProperties(columnwidth="{}cm".format(size)))
                self.doc.automaticstyles.addElement(sc)

            #Cell header style
            sch=Style(name="Table.HeaderCell", family="table-cell")
            sch.addElement(TableCellProperties(backgroundcolor="#999999",  border="0.05pt solid #000000", padding="0.15cm"))
            self.doc.automaticstyles.addElement(sch)        

            #Cell normal
            sch=Style(name="Table.Cell", family="table-cell")
            sch.addElement(TableCellProperties(border="0.05pt solid #000000", padding="0.1cm"))
            self.doc.automaticstyles.addElement(sch)

            #TAble contents style
            s= Style(name="Table.Heading.Font{}".format(fontsize), family="paragraph" )
            s.addElement(TextProperties(attributes={'fontsize':"{}pt".format(fontsize),  'fontweight':"bold"}))
            s.addElement(ParagraphProperties(attributes={'textalign':'center', }))
            self.doc.styles.addElement(s)

            s = Style(name="Table.Contents.Font{}".format(fontsize), family="paragraph")
            s.addElement(TextProperties(attributes={'fontsize':"{}pt".format(fontsize), }))
            s.addElement(ParagraphProperties(attributes={'textalign':'justify', }))
            self.doc.styles.addElement(s)

            s = Style(name="Table.ContentsRight.Font{}".format(fontsize), family="paragraph")
            s.addElement(TextProperties(attributes={'fontsize':"{}pt".format(fontsize), }))
            s.addElement(ParagraphProperties(attributes={'textalign':'end', }))
            self.doc.styles.addElement(s)

            s = Style(name="Table.ContentsRight.FontRed{}".format(fontsize), family="paragraph")
            s.addElement(TextProperties(attributes={'fontsize':"{}pt".format(fontsize), 'color':'#ff0000' }))
            s.addElement(ParagraphProperties(attributes={'textalign':'end', }))
            self.doc.styles.addElement(s)

        ## Generate a TableCell guessing style, setting color if number is negative and setting alignment
        ## @param o Object can be str, Decimal, int, float, datetime, date, Currency, Percentage
        ## @param fontsize Integer with the size of the font
        ## @return TableCell
        def addTableCell(o, fontsize):
            tc = TableCell(stylename="Table.Cell")
            #Parses orientation
            p = P(stylename="Table.ContentsRight.Font{}".format(fontsize))
            s=Span(text=str(o))
            if o.__class__.__name__  == "datetime":
                p = P(stylename="Table.Contents.Font{}".format(fontsize))
                s=Span(text=dtnaive2string(o, "%Y-%m-%d %H:%M:%S"))
            if o.__class__.__name__ in ("str", "date" ):
                p = P(stylename="Table.Contents.Font{}".format(fontsize))
                s=Span(text=str(o))
            elif o.__class__.__name__ in ("Currency", "Percentage", "Money"):
                if o.isLTZero():
                    p = P(stylename="Table.ContentsRight.FontRed{}".format(fontsize))
                s=Span(text=o.string())
            elif o.__class__ in ("int", "Decimal",  "float"):
                if o<0:
                    p = P(stylename="Table.ContentsRight.FontRed{}".format(fontsize))
            p.addElement(s)
            tc.addElement(p)
            return tc

        ######################################
        self.seqTables=self.seqTables+1
        generate_table_styles()
        #Table columns
        name=name if name!=None else "Table.{}".format(self.seqTables)
        table = Table(name=name, stylename="Table.Size{}".format(sum(sizes)))
        for size in sizes:
            table.addElement(TableColumn(stylename="Table.Column.Size{}".format(size)))
        #Header rows
        headerrow=TableHeaderRows()
        tablerow=TableRow()
        headerrow.addElement(tablerow)
        if hh is not None:
            for i, head in enumerate(hh):
                p=P(stylename="Table.Heading.Font{}".format(fontsize), text=head)
                tablecell=TableCell(stylename="Table.HeaderCell")
                tablecell.addElement(p)
                tablerow.addElement(tablecell)
        table.addElement(headerrow)

        #Data rows
        for row in data:
            tr = TableRow()
            table.addElement(tr)
            for i, o in enumerate(row):
                tr.addElement(addTableCell(o, fontsize))
        return self.insertInCursor(table, after)

    ## Adds a paragraph with illustration style, with a list with image keys. All images will have the same width and height
    ## @param image_key_list List with image keys to add to the paragraph
    ## @param width Integer or float with number of width centimeters of the image in the document
    ## @param height Integer or float with number of height centimeters of the image in the document
    ## @param name String With the root name of the images in the paragraph. The root will be appended with a sequential number
    ## @param after True: insert after self.cursor element. False: insert before self.cursor element. None: Just return element
    def illustration(self, image_key_list, width, height, name, after=True):
        p = P(stylename="Illustration")
        for i, image_key in enumerate(image_key_list):
            if name==None:
                self.seqFrames=self.seqFrames+1
                n="Frame.{}".format(self.seqFrames)
            else:
                n="{}.{}".format(name, i)
            p.addElement(self.image(image_key, width, height, n))
        self.insertInCursor(p, after=True)

## Class with starndard.odt template
class ODT_Standard(ODT):
    def __init__(self, filename, language="es", country="es"):
        template=resource_filename("officegenerator","templates/odt/standard.odt")
        ODT.__init__(self, filename, template, language, country)

    ## PH and PV styles are deffined in standard.odtg
    def pageBreak(self,  horizontal=False, after=True):    
        if horizontal==True:
            p=P(stylename="PH")
        else:
            p=P(stylename="PV")
        return self.insertInCursor(p, after)

## Class with starndard.odt template
class ODT_Manual_Styles(ODT):
    def __init__(self, filename, language="es", country="es"):
        ODT.__init__(self, filename, None, language, country)
        self.load_predefined_styles()

    ## Creates a text header
    ## @param text String with the header string
    ## @Level Integer Level of the header
    def header(self, text, level, after=True):
        h=H(outlinelevel=level, stylename="Heading{}".format(level), text=text)
        return self.insertInCursor(h, after)
    ## Loads predefined styles. If you want to change them you need to make a class with ODT as its parent and override this method or to load a template
    def load_predefined_styles(self):
        def stylePage():
            pagelayout=PageLayout(name="PageLayout")
            plp=PageLayoutProperties(pagewidth="21cm",  pageheight="29.7cm",  margintop="2cm",  marginright="2cm",  marginleft="2cm",  marginbottom="2cm")
            fs=FooterStyle()
            hfp=HeaderFooterProperties(margintop="0.5cm")
            fs.addElement(hfp)
            pagelayout.addElement(plp)
            pagelayout.addElement(fs)
            self.doc.automaticstyles.addElement(pagelayout)

        def styleParagraphs():
            #Pagebreak styles horizontal y vertical        
            s = Style(name="PH", family="paragraph",  parentstylename="Standard", masterpagename="Landscape")
            s.addElement(ParagraphProperties(pagenumber="auto"))
            self.doc.styles.addElement(s)
            s = Style(name="PV", family="paragraph",  parentstylename="Standard", masterpagename="Standard")
            s.addElement(ParagraphProperties(pagenumber="auto"))
            self.doc.styles.addElement(s)

            standard= Style(name="Standard", family="paragraph",  autoupdate="true")
            standard.addElement(ParagraphProperties(attributes={"margintop":"0.2cm", "textalign":"justify", "marginbottom":"0.2cm", "textindent":"1.5cm"}))
            standard.addElement(TextProperties(attributes={"fontsize": "12pt", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(standard)   

            ImageCenter= Style(name="Illustration", family="paragraph",  autoupdate="true")
            ImageCenter.addElement(ParagraphProperties(attributes={"margintop":"1cm", "textalign":"center", "marginbottom":"1cm"}))
            ImageCenter.addElement(TextProperties(fontsize="12pt", fontstyle="italic", country=self.country, language=self.language))
            self.doc.styles.addElement(ImageCenter)

            standardCenter= Style(name="StandardCenter", family="paragraph",  autoupdate="true")
            standardCenter.addElement(ParagraphProperties(attributes={"margintop":"0.2cm", "textalign":"center", "marginbottom":"0.2cm", "textindent":"0cm"}))
            standardCenter.addElement(TextProperties(attributes={"fontsize": "12pt", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(standardCenter)

            standardRight= Style(name="StandardRight", family="paragraph",  autoupdate="true")
            standardRight.addElement(ParagraphProperties(attributes={"margintop":"0.2cm", "textalign":"right", "marginbottom":"0.2cm", "textindent":"0cm"}))
            standardRight.addElement(TextProperties(attributes={"fontsize": "12pt", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(standardRight)



            letra12= Style(name="Bold12Underline", family="paragraph",  autoupdate="true")
            letra12.addElement(ParagraphProperties(attributes={"margintop":"0.4cm", "textalign":"justify", "marginbottom":"0.4cm", "textindent":"1.5cm"}))
            letra12.addElement(TextProperties(attributes={"fontsize": "12pt", "fontweight": "bold", "country": self.country, "language": self.language, 'textunderlinecolor':"font-color", 'textunderlinewidth':"auto", 'textunderlinestyle':"solid"}))
            self.doc.styles.addElement(letra12)


            letra18= Style(name="Bold18Center", family="paragraph",  autoupdate="true")
            letra18.addElement(ParagraphProperties(attributes={"margintop":"0.2cm", "textalign":"center", "marginbottom":"0.2cm", "textindent":"0cm"}))
            letra18.addElement(TextProperties(attributes={"fontsize": "18pt", "fontweight": "bold", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(letra18)

            letra16= Style(name="Bold16Center", family="paragraph",  autoupdate="true")
            letra16.addElement(ParagraphProperties(attributes={"margintop":"0.2cm", "textalign":"center", "marginbottom":"0.2cm", "textindent":"0cm"}))
            letra16.addElement(TextProperties(attributes={"fontsize": "16pt", "fontweight": "bold", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(letra16)

            letra14= Style(name="Bold14Center", family="paragraph",  autoupdate="true")
            letra14.addElement(ParagraphProperties(attributes={"margintop":"0.2cm", "textalign":"center", "marginbottom":"0.2cm", "textindent":"0cm"}))
            letra14.addElement(TextProperties(attributes={"fontsize": "14pt", "fontweight": "bold", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(letra14)

            letra12= Style(name="Bold12Center", family="paragraph",  autoupdate="true")
            letra12.addElement(ParagraphProperties(attributes={"margintop":"0.2cm", "textalign":"center", "marginbottom":"0.2cm", "textindent":"0cm"}))
            letra12.addElement(TextProperties(attributes={"fontsize": "12pt", "fontweight": "bold", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(letra12)

        def styleHeaders():
            s = Style(name="Title", family="paragraph",  autoupdate="true", defaultoutlinelevel="0")
            s.addElement(ParagraphProperties(attributes={"margintop":"0.7cm", "textalign":"center", "marginbottom":"0.7cm"}))
            s.addElement(TextProperties(attributes={"fontsize": "18pt", "fontweight": "bold", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(s)

            s = Style(name="Subtitle", family="paragraph",  autoupdate="true", defaultoutlinelevel="0")
            s.addElement(ParagraphProperties(attributes={"margintop":"0cm", "textalign":"center", "marginbottom":"0.7cm"}))
            s.addElement(TextProperties(attributes={"fontsize": "15pt", "fontstyle": "italic", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(s)

            h1style = Style(name="Heading1", family="paragraph",  autoupdate="true", defaultoutlinelevel="1")
            h1style.addElement(ParagraphProperties(attributes={"margintop":"0.6cm", "textalign":"justify", "marginbottom":"0.3cm"}))
            h1style.addElement(TextProperties(attributes={"fontsize": "15pt", "fontweight": "bold", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(h1style)

            h2style = Style(name="Heading2", family="paragraph",  autoupdate="true", defaultoutlinelevel="2")
            h2style.addElement(ParagraphProperties(attributes={"margintop":"0.5cm", "textalign":"justify", "marginbottom":"0.25cm"}))
            h2style.addElement(TextProperties(attributes={"fontsize": "14pt", "fontweight": "bold", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(h2style)

            out=OutlineStyle(name="Outline")
            outl=OutlineLevelStyle(level=1, numformat="1", numsuffix="  ")
            out.addElement(outl)
            outl=OutlineLevelStyle(level=2, displaylevels="2", numformat="1", numsuffix="  ")
            out.addElement(outl)
            self.doc.styles.addElement(out)

        def styleList():
            liststandard= Style(name="ListStandard", family="paragraph",  autoupdate="true")
            liststandard.addElement(ParagraphProperties(attributes={"margintop":"0.1cm", "textalign":"justify", "marginbottom":"0.1cm", "textindent":"0cm"}))
            liststandard.addElement(TextProperties(attributes={"fontsize": "12pt", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(liststandard)
            
            # For Bulleted list
            bulletedliststyle = ListStyle(name="BulletList1")
            bulletlistproperty = ListLevelStyleBullet(level="1", bulletchar=u"•")
            bulletlistproperty.addElement(ListLevelProperties( minlabelwidth="1cm"))
            bulletedliststyle.addElement(bulletlistproperty)
            self.doc.styles.addElement(bulletedliststyle)
            # For Bulleted list
            bulletedliststyle = ListStyle(name="BulletList2")
            bulletlistproperty = ListLevelStyleBullet(level="2", bulletchar=u"#")
            bulletlistproperty.addElement(ListLevelProperties( minlabelwidth="1cm"))
            bulletedliststyle.addElement(bulletlistproperty)
            self.doc.styles.addElement(bulletedliststyle)
            # For Bulleted list
            bulletedliststyle = ListStyle(name="BulletList2")
            bulletlistproperty = ListLevelStyleBullet(level="2", bulletchar=u"·")
            bulletlistproperty.addElement(ListLevelProperties( minlabelwidth="1cm"))
            bulletedliststyle.addElement(bulletlistproperty)
            self.doc.styles.addElement(bulletedliststyle)

            # For numbered list
            numberedliststyle = ListStyle(name="NumberedList")
            numberedlistproperty = ListLevelStyleNumber(level="1", numsuffix=".", startvalue=1)
            numberedlistproperty.addElement(ListLevelProperties(minlabelwidth="1cm"))
            numberedliststyle.addElement(numberedlistproperty)
            self.doc.styles.addElement(numberedliststyle)
        
        def styleFooter():
            s= Style(name="Footer", family="paragraph",  autoupdate="true")
            s.addElement(ParagraphProperties(attributes={"margintop":"0cm", "textalign":"center", "marginbottom":"0cm", "textindent":"0cm"}))
            s.addElement(TextProperties(attributes={"fontsize": "9pt", "country": self.country, "language": self.language}))
            self.doc.styles.addElement(s)

        def styleMasterPage():
            foot=MasterPage(name="Standard", pagelayoutname="PageLayout")
            footer=Footer()
            p1=P(stylename="Footer",  text="Página ")
            number=PageNumber(selectpage="current", numformat="1")
            p2=Span(stylename="Footer",  text=" de  ")
            count=PageCount(selectpage="current", numformat="1")
            p1.addElement(number)
            p1.addElement(p2)
            p1.addElement(count)      
            footer.addElement(p1)   
            foot.addElement(footer)
            self.doc.masterstyles.addElement(foot)
        #######################################                
        stylePage()
        styleParagraphs()
        styleFooter()
        styleMasterPage()
        styleHeaders()
        styleList()

    ## This function always use manual Bulletlist and liststandard styles by coherence with manual styles
    def list(self, arr, list_style=None, paragraph_style=None,  after=True):
        return ODT.list(self, arr, "BulletList", "ListStandard", after)

    def pageBreak(self,  horizontal=False, after=True):    
        p=P(stylename="PageBreak")#Is an automatic style
        self.doc.text.addElement(p)
        if horizontal==True:
            p=P(stylename="PH")
        else:
            p=P(stylename="PV")
        return self.insertInCursor(p, after)

## Manage Odf Cells
class OdfCell:
    def __init__(self, coord,  object, style):
        self.coord=Coord.assertCoord(coord)
        self.object=object
        self.style=style

        self.spannedColumns=1
        self.spannedRows=1
        self.comment=None

    def __repr__(self):
        return "OdfCell <{}{}>".format(self.coord.letter, self.coord.number)

    def generate(self):
        if self.object==None:
            self.object=" - "
        if self.object.__class__.__name__ in ["Currency", "Money"]:
            odfpycell = TableCell(valuetype="currency", currency=self.object.currency, value=self.object.amount, stylename=self.style)
        elif self.object.__class__.__name__=="Percentage":
            odfpycell = TableCell(valuetype="percentage", value=self.object.value, stylename=self.style)
        elif self.object.__class__.__name__=="datetime":
            odfpycell = TableCell(valuetype="date", datevalue=self.object.strftime("%Y-%m-%dT%H:%M:%S"), stylename=self.style)
        elif self.object.__class__.__name__=="time":
            odfpycell = TableCell(valuetype="time", timevalue=self.object.strftime("PT%HH%MM%SS"), stylename=self.style)
        elif self.object.__class__.__name__=="date":
            odfpycell = TableCell(valuetype="date", datevalue=str(self.object), stylename=self.style)
        elif self.object.__class__.__name__ in ("Decimal", "float", "int"):
            odfpycell= TableCell(valuetype="float", value=self.object,  stylename=self.style)
        elif self.object.__class__.__name__ =="bool":
            odfpycell= TableCell(valuetype="boolean", booleanvalue=self.object, stylename=self.style)
        elif self.object.__class__.__name__ =="Formula":
            odfpycell = TableCell(formula="of:"+self.object.string_formula,  stylename=self.style)
        else:#strings
            odfpycell = TableCell(valuetype="string", value=self.object,  stylename=self.style)
            odfpycell.addElement(P(text = self.object))

        if self.spannedRows!=1 or self.spannedColumns!=1:
            odfpycell.setAttribute("numberrowsspanned", str(self.spannedRows))
            odfpycell.setAttribute("numbercolumnsspanned", str(self.spannedColumns))
        if self.comment!=None:
            a=Annotation(textstylename="Right")
            d=Date()
            d.addText(datetime.now().strftime("%Y-%m-%dT%H:%M:%S"))
            a.addElement(d)
            a.addElement(P(stylename="Right", text=self.comment))
            odfpycell.addElement(a)
        return odfpycell

    ## Manage cell spannning 
    ##
    ## Siempre es de izquierda a derecha y de arriba a abajo. Si es 1 no hay spanning
    def setSpanning(self, columns, rows):
        self.spannedColumns=columns
        self.spannedRows=rows

    def setComment(self, comment):
        self.comment=comment


## Class to create a sheet in a ods document
## By default cursor position and split position is set to "A1" cell
class OdfSheet:
    def __init__(self, doc,  title):
        self.doc=doc
        self.title=title
        self.widths=None
        self.arr=[]
        self.freezeAndSelect("A1", "A1", "A1")#Default values

    ## Freeze panels in a sheet and sets the selected cell
    ##        split/freeze vertical (0|1|2) - 1 = split ; 2 = freeze
    ##          ##split/freeze horizontal (0|1|2) - 1 = split ; 2 = freeze
    ##      vertical position = in cell if fixed, in screen unit if frozen
    ##      horizontal position = in cell if fixed, in screen unit if frozen
    ##      active zone in the splitted|frozen sheet (0..3 from let to right, top
    ##  to bottom)
    ##  #   COMPROBADO CON ODF2XML
    ##  B1: 
    ##                <config:config-item config:name="HorizontalSplitMode" config:type="short">2</config:config-item>
    ##                <config:config-item config:name="VerticalSplitMode" config:type="short">0</config:config-item>
    ##                <config:config-item config:name="HorizontalSplitPosition" config:type="int">1</config:config-item>
    ##                <config:config-item config:name="VerticalSplitPosition" config:type="int">0</config:config-item>
    ##                <config:config-item config:name="ActiveSplitRange" config:type="short">3</config:config-item>
    ##                <config:config-item config:name="PositionLeft" config:type="int">0</config:config-item>
    ##                <config:config-item config:name="PositionRight" config:type="int">1</config:config-item>
    ##                <config:config-item config:name="PositionTop" config:type="int">0</config:config-item>
    ##                <config:config-item config:name="PositionBottom" config:type="int">0</config:config-item>
    ## @param freeze_coord, Cell where panels are frrozen. Can be a string or a Coord object.
    ## @param selected_coord. Cell selected opening sheet. Can be a string or a Coord object.
    ## @param topLeftCell, topleftcell to show in sheet after opening. Can be a string or a Coord object.
    def freezeAndSelect(self, freeze_coord, selected_coord=None, topleftcell_coord=None):
        if selected_coord is None:
            selected_coord=Coord(self.lastColumn()+self.lastRow())
        
        if topleftcell_coord is None:
            topleftcell_coord=topLeftCellNone(freeze_coord, selected_coord)
        # Creates Coord objects
        freeze_coord=Coord.assertCoord(freeze_coord)
        selected_coord=Coord.assertCoord(selected_coord)
        topleftcell_coord=Coord.assertCoord(topleftcell_coord)

        #Sets cursor position
        self.cursorPositionX=selected_coord.letterIndex()
        self.cursorPositionY=selected_coord.numberIndex()
    
        #Sets freeze position and modes
        self.horizontalSplitPosition=str(freeze_coord.letterIndex())
        self.verticalSplitPosition=str(freeze_coord.numberIndex())
        self.horizontalSplitMode="0" if self.horizontalSplitPosition=="0" else "2"
        self.verticalSplitMode="0" if self.verticalSplitPosition=="0" else "2"
        
        #Sets active split range and top left cell position
        if self.horizontalSplitPosition!="0" and self.verticalSplitPosition=="0":#C1 WORKS
            self.activeSplitRange="3"
            self.positionTop="0"
            self.positionBottom=str(topleftcell_coord.numberIndex())
            self.positionLeft="0"
            self.positionRight=str(topleftcell_coord.letterIndex())
        elif self.horizontalSplitPosition=="0" and self.verticalSplitPosition!="0":#A3 WORKS
            self.activeSplitRange="2"
            self.positionTop="0"
            self.positionBottom=str(topleftcell_coord.numberIndex())
            self.positionLeft=str(topleftcell_coord.letterIndex())
            self.positionRight="0"
        elif self.horizontalSplitPosition!="0" and self.verticalSplitPosition!="0": # C3  WORKS
            self.activeSplitRange="3"
            self.positionTop="0"
            self.positionBottom=str(topleftcell_coord.numberIndex())
            self.positionLeft="0"
            self.positionRight=str(topleftcell_coord.letterIndex())
        else:#A1 WORKS
            self.activeSplitRange="2"
            self.positionTop="0"
            self.positionBottom=str(topleftcell_coord.numberIndex())
            self.positionLeft=str(topleftcell_coord.letterIndex())
            self.positionRight=str(topleftcell_coord.letterIndex())

    ## Sets a comment in the givven cell
    ## @param coord can be Coord o Coord.string()
    ## @param comment String to insert as a comment in the cell
    def setComment(self, coord, comment):
        coord=Coord.assertCoord(coord)
        c=self.getCell(coord)
        c.setComment(comment)


    def setColumnsWidth(self, widths, unit="pt"):
        """
            widths is an int array
            id es el id del sheet de python
        """
        for w in widths:
            s=Style(name="{}_{}".format(id(self), w), family="table-column")
            s.addElement(TableColumnProperties(columnwidth="{}{}".format(w, unit)))
            self.doc.automaticstyles.addElement(s)   
        self.widths=widths
        
    ## Returns the last letter used in the sheet . Returns a string with the letter name of the column
    def lastColumn(self):
        return number2column(self.columns())

    ## Returns the last  row name used
    ## @return string row name
    def lastRow(self):
        return  number2row(self.rows())

    def addCell(self, cell): 
        self.arr.append(cell)
        
    ## Returns the cell in the Coord
    ## @return OdfCell
    def getCell(self, coord):
        coord=Coord.assertCoord(coord)
        for c in self.arr:
            if c.coord==coord:
                return c
        return None
            
    def getCellValue(self, coord):
        cell=self.getCell(coord)
        if cell==None:
            return None
        else:
            return self.getCell(coord).object
    
    ## Returns true if value is a string beginning with = or +
    ## @param value must be a string
    ## @return boolean
    def isFormula(self, value):
        if len(value)>0 and value[0] in ["=", "+"]:
            return True
        return False

    ## Adds a cell to the sheet using its coord, an object and a color or a style
    ## @param coord Coord where the cell is going to be created
    ## @param result Object to add to the Cell. Can be int, str, float, datetime, date, Currency, Percentage, Decimal, None (will be converted to " - ")
    ## @param color_or_style String with a color: Normal, White, Yellow, Orange, Blue, Red, GrayLight, GrayDark. Or a style WhiteInteger, YellowLeft, OrangeCenter, OrangeEUR, RedPercentage...
    def add(self, coord, result, color_or_style="Normal"):
        coord=Coord.assertCoord(coord)
        if isFormula(result):
            result=Formula(result)
        

        if result.__class__ in (list,):#Una lista
            for i,row in enumerate(result):
                if row.__class__ in (list, ):#Una lista de varias columnas
                    for j, column in enumerate(row):
                        style=guess_ods_style(color_or_style, result[i][j])
                        self.addCell(OdfCell(Coord(coord.string()).addColumn(j).addRow(i), result[i][j], style))
                else: #Any value not list if row.__class__ in (int, str, float, datetime,  date, Currency, Percentage,  Decimal, bool):#Una lista de una columna
                    style=guess_ods_style(color_or_style, result[i])
                    self.addCell(OdfCell(Coord(coord.string()).addRow(i), result[i], style))
        else: #Any value not list#result.__class__ in (str, int, float, datetime, date,  Currency, Percentage, Decimal,bool):#Un solo valor
            style=guess_ods_style(color_or_style, result)  
            self.addCell(OdfCell(coord, result, style))

    ## Adds a cell to self.arr with merge, content and style information
    ## @param range Range
    def addMerged(self, range, result, style):
        range=Range.assertRange(range)
        self.add(range.start, result, style)      
        c=self.getCell(range.start)
        c.setSpanning(range.numColumns(), range.numRows())
            
    ## @param cood Coord from we are going to add totals
    ## @param list_of_totals List with strings or keys. Example: ["Total", "#SUM", "#AVG"]...
    ## @param list_of_styles List with string styles or None. If none tries to guest from top column object. List example: ["GrayLightPercentage", "GrayLightInteger"]
    ## @param string with the row where th3e total begins
    ## @param string with the rew where the formula ends. If None it's a coord.row -1
    def addTotalsHorizontal(self, coord, list_of_totals, list_of_styles=None, row_from="2", row_to=None):
        coord=Coord.assertCoord(coord)
        for letter, total in enumerate(list_of_totals):
            coord_total=coord.addColumnCopy(letter)
            coord_total_from=Coord(coord_total.letter+row_from)
            if row_to is None:
                coord_total_to=coord_total.addRowCopy(-1)# row above
            else:
                coord_total=Coord(coord_total.letter+row_to)

            if list_of_styles is None:
                style=guess_ods_style("GrayLight", self.getCellValue(coord_total_from))
            else:
                style=list_of_styles[letter]

            self.add(coord_total, generate_formula_total_string(total, coord_total_from, coord_total_to),  style)
            
    ## @param cood Coord from we are going to add totals
    ## @param list_of_totals List with strings or keys. Example: ["Total", "#SUM", "#AVG"]...
    ## @param list_of_styles List with string styles or None. If none tries to guest from top column object. List example: ["GrayLightPercentage", "GrayLightInteger"]
    ## @param string with the row where th3e total begins
    ## @param string with the rew where the formula ends. If None it's a coord.row -1
    def addTotalsVertical(self, coord, list_of_totals, list_of_styles=None, column_from="B", column_to=None):
        coord=Coord.assertCoord(coord)
        for i, total in enumerate(list_of_totals):
            coord_total=coord.addRowCopy(i)
            coord_total_from=Coord(column_from+coord_total.number)
            if column_to is None:
                coord_total_to=coord_total.addColumnCopy(-1)# row above
            else:
                coord_total=Coord(column_to+coord_total.number)

            if list_of_styles is None:
                style=guess_ods_style("GrayLight", self.getCellValue(coord_total_from))
            else:
                style=list_of_styles[i]

            self.add(coord_total, generate_formula_total_string(total, coord_total_from, coord_total_to),  style)

    ## Generates the sheet in self.doc Opendocument varianble
    def generate(self, ods):
        # Start the table
        columns=self.columns()
        rows=self.rows()
        grid=[[None for x in range(columns)] for y in range(rows)]
        for cell in self.arr:
            grid[cell.coord.numberIndex()][cell.coord.letterIndex()]=cell

        table = Table(name=self.title)
        for c in range(columns):#Create columns
            try:
                tc=TableColumn(stylename="{}_{}".format(id(self), self.widths[c]))
            except:
                tc=TableColumn()
            table.addElement(tc)
        for j in range(rows):#Crreate rows
            tr = TableRow()
            table.addElement(tr)
            for i in range(columns):#Create cells
                cell=grid[j][i]
                if cell is not None:
                    tr.addElement(cell.generate())
                else:
                    tr.addElement(TableCell())
        ods.doc.spreadsheet.addElement(table)

    ## Gets the number of columns that are used in the sheet
    ## @return int
    def columns(self):
        r=0
        for cell in self.arr:
            column=cell.coord.letterPosition()
            if column>r:
                r=column
        return r

    ## Return the number of rows that are used in the cell.
    def rows(self):
        r=0
        for cell in self.arr:
            column=cell.coord.numberPosition()
            if column>r:
                r=column
        return r


## Abstract class with common functions to generate Libreoffice Calc ODS files
class ODS(ODF):
    def __init__(self, filename):
        ODF.__init__(self, filename)
        self.doc=OpenDocumentSpreadsheet()
        self.sheets=[]
        self.activeSheet=None

    def createSheet(self, title):
        s=OdfSheet(self.doc, title)
        self.sheets.append(s)
        return s

    def getActiveSheet(self):
        if self.activeSheet==None:
            return self.sheets[0].title
        return self.activeSheet

    ## Save ODS file
    ## @param filename str or None. If filename is given, file is saved with a different name to the constructor one
    def save(self, filename=None):
        #config settings information
        a=ConfigItemSet(name="ooo:view-settings")
        aa=ConfigItem(type="int", name="VisibleAreaTop")
        aa.addText("0")
        a.addElement(aa)
        aa=ConfigItem(type="int", name="VisibleAreaLeft")
        aa.addText("0")
        a.addElement(aa)
        b=ConfigItemMapIndexed(name="Views")
        c=ConfigItemMapEntry()
        d=ConfigItem(name="ViewId", type="string")
        d.addText("view1")#value="view1"
        e=ConfigItemMapNamed(name="Tables")
        for sheet in self.sheets:
            f=ConfigItemMapEntry(name=sheet.title)
            g=ConfigItem(type="int", name="CursorPositionX")
            g.addText(sheet.cursorPositionX)
            f.addElement(g)
            g=ConfigItem(type="int", name="CursorPositionY")
            g.addText(sheet.cursorPositionY)
            f.addElement(g)
            g=ConfigItem(type="int", name="HorizontalSplitPosition")
            g.addText(sheet.horizontalSplitPosition)
            f.addElement(g)
            g=ConfigItem(type="int", name="VerticalSplitPosition")
            g.addText(sheet.verticalSplitPosition)
            f.addElement(g)
            g=ConfigItem(type="short", name="HorizontalSplitMode")
            g.addText(sheet.horizontalSplitMode)
            f.addElement(g)
            g=ConfigItem(type="short", name="VerticalSplitMode")
            g.addText(sheet.verticalSplitMode)
            f.addElement(g)
            g=ConfigItem(type="short", name="ActiveSplitRange")
            g.addText(sheet.activeSplitRange)
            f.addElement(g)
            g=ConfigItem(type="int", name="PositionLeft")
            g.addText(sheet.positionLeft)
            f.addElement(g)
            g=ConfigItem(type="int", name="PositionRight")
            g.addText(sheet.positionRight)
            f.addElement(g)
            g=ConfigItem(type="int", name="PositionTop")
            g.addText(sheet.positionTop)
            f.addElement(g)
            g=ConfigItem(type="int", name="PositionBottom")
            g.addText(sheet.positionBottom)
            f.addElement(g)
            e.addElement(f)
            
        a.addElement(b)
        b.addElement(c)
        c.addElement(d)
        c.addElement(e)

        h=ConfigItem(type="string", name="ActiveTable")
        h.addText(self.getActiveSheet())
        c.addElement(h)
        self.doc.settings.addElement(a)
        
        for sheet in self.sheets:
            sheet.generate(self)
        
        if  filename==None:
            filename=self.filename

        if path.dirname(filename)!="":
            makedirs(path.dirname(filename), exist_ok=True)
        self.doc.save(filename)

    def setActiveSheet(self, value):
        """value is OdfSheet"""
        self.activeSheet=value.title

class ODSStyleCurrency:
    def __init__(self, name, symbol):
        self.name=name
        self.symbol=symbol
        
    ## Generate currency styles
    def generate_ods_styles(self, doc):
        # Create the styles for $AUD format currency values
        ns1 = CurrencyStyle(name=self.name + "Black", volatile="true")
        ns1.addElement(Number(decimalplaces="2", minintegerdigits="1", grouping="true"))
        ns1.addElement(CurrencySymbol(language="es", country="ES", text=" "+ self.symbol))
        doc.styles.addElement(ns1)

        # Create the main style.
        ns2 = CurrencyStyle(name=self.name)
        ns2.addElement(TextProperties(color="#ff0000"))
        ns2.addElement(Text(text="-"))
        ns2.addElement(Number(decimalplaces="2", minintegerdigits="1", grouping="true"))
        ns2.addElement(CurrencySymbol(language="es", country="ES", text=" "+ self.symbol))
        ns2.addElement(Map(condition="value()>=0", applystylename=self.name + "Black"))
        doc.styles.addElement(ns2)
        

## Class to manage color in Libreoffice Calc ODS. It generates all needed ODS styles needed for that color
class ODSStyleColor:
    def __init__(self, name, rgb, bold):
        self.name=name
        self.rgb=rgb
        self.bold=bold

    def generate_ods_styles(self, doc, currencymanager):
        hs=Style(name=self.name + "Center", family="table-cell")
        hs.addElement(TableCellProperties(backgroundcolor=self.rgb, border="0.06pt solid #000000", verticalalign="middle", textalignsource="fix"))
        if self.bold==True:
            hs.addElement(TextProperties( fontweight="bold"))
        hs.addElement(ParagraphProperties(textalign="center"))
        doc.styles.addElement(hs)

        hs=Style(name=self.name + "Left", family="table-cell")
        hs.addElement(TableCellProperties(backgroundcolor=self.rgb, border="0.06pt solid #000000", verticalalign="middle", textalignsource="fix"))
        if self.bold==True:
            hs.addElement(TextProperties( fontweight="bold"))
        hs.addElement(ParagraphProperties(textalign="left"))
        doc.styles.addElement(hs)

        hs=Style(name=self.name + "Right", family="table-cell")
        hs.addElement(TableCellProperties(backgroundcolor=self.rgb, border="0.06pt solid #000000", verticalalign="middle", textalignsource="fix"))
        if self.bold==True:
            hs.addElement(TextProperties( fontweight="bold"))
        hs.addElement(ParagraphProperties(textalign="end"))
        doc.styles.addElement(hs)

        for currency in currencymanager.arr:
            moneycontents = Style(name=self.name+currency.name, family="table-cell",  datastylename=currency.name ,parentstylename=self.name+"Right")
            doc.styles.addElement(moneycontents)

        pourcent = Style(name=self.name+'Percentage', family='table-cell', datastylename='Percentage',parentstylename=self.name+"Right")
        doc.styles.addElement(pourcent)

        dt = Style(name=self.name+"Datetime", datastylename="Datetime",parentstylename=self.name+"Left", family="table-cell")
        doc.styles.addElement(dt)

        dat = Style(name=self.name+"Date", datastylename="Date",parentstylename=self.name+"Left", family="table-cell")
        doc.styles.addElement(dat)
        
        time = Style(name=self.name+"Time", datastylename="Time",parentstylename=self.name+"Left", family="table-cell")
        doc.styles.addElement(time)

        integer = Style(name=self.name+"Integer", family="table-cell",  datastylename="Integer",parentstylename=self.name+"Right")
        doc.styles.addElement(integer)
        
        boolean = Style(name=self.name+"Boolean", family="table-cell",  datastylename="Boolean",parentstylename=self.name+"Right")
        doc.styles.addElement(boolean)

        decimal2= Style(name=self.name+"Decimal2", family="table-cell",  datastylename="Decimal2",parentstylename=self.name+"Right")
        doc.styles.addElement(decimal2)

        decimal6= Style(name=self.name+"Decimal6", family="table-cell",  datastylename="Decimal6",parentstylename=self.name+"Right")
        doc.styles.addElement(decimal6)


## Class to Mange ODSStyleColors in libodfgenerator
class ODSStyleCurrencyManager:
    def __init__(self):
        self.arr=[]

    def append(self, o):
        self.arr.append(o)

    def generate_ods_styles(self, doc):
        for o in self.arr:
            o.generate_ods_styles(doc)

## Class to Mange ODSStyleColors in libodfgenerator
class ODSStyleColorManager:
    def __init__(self):
        self.arr=[]

    def append(self, o):
        self.arr.append(o)

    ## Generate common styles used in all color styles
    def __generate_ods_common_styles(self, doc):
        #Percentage
        nonze = PercentageStyle(name='PercentageBlack')
        nonze.addElement(TextProperties(color="#000000"))
        nonze.addElement(Number(decimalplaces='2', minintegerdigits='1'))
        nonze.addElement(Text(text=' %'))
        doc.styles.addElement(nonze)
        
        nonze2 = PercentageStyle(name='Percentage')
        nonze2.addElement(TextProperties(color="#ff0000"))
        nonze2.addElement(Text(text="-"))
        nonze2.addElement(Number(decimalplaces='2', minintegerdigits='1'))
        nonze2.addElement(Text(text=' %'))
        nonze2.addElement(Map(condition="value()>=0", applystylename="PercentageBlack"))
        doc.styles.addElement(nonze2)
        
        # Datetimes
        date_style = DateStyle(name="Datetime") #, language="lv", country="LV")
        date_style.addElement(Year(style="long"))
        date_style.addElement(Text(text="-"))
        date_style.addElement(Month(style="long"))
        date_style.addElement(Text(text="-"))
        date_style.addElement(Day(style="long"))
        date_style.addElement(Text(text=" "))
        date_style.addElement(Hours(style="long"))
        date_style.addElement(Text(text=":"))
        date_style.addElement(Minutes(style="long"))
        date_style.addElement(Text(text=":"))
        date_style.addElement(Seconds(style="long"))
        doc.styles.addElement(date_style)#NO SERIA ESTE UN AUTOMATICO????        
        
        # Time
        time_style = TimeStyle(name="Time") #, language="lv", country="LV")
        time_style.addElement(Hours(style="long"))
        time_style.addElement(Text(text=":"))
        time_style.addElement(Minutes(style="long"))
        time_style.addElement(Text(text=":"))
        time_style.addElement(Seconds(style="long"))
        doc.styles.addElement(time_style)#NO SERIA ESTE UN AUTOMATICO????

        #Date
        date_style = DateStyle(name="Date") #, language="lv", country="LV")
        date_style.addElement(Year(style="long"))
        date_style.addElement(Text(text="-"))
        date_style.addElement(Month(style="long"))
        date_style.addElement(Text(text="-"))
        date_style.addElement(Day(style="long"))
        doc.styles.addElement(date_style)

        #Integer
        ns1 = NumberStyle(name="IntegerBlack", volatile="true")
        ns1.addElement(Number(decimalplaces="0", minintegerdigits="1", grouping="true"))
        doc.styles.addElement(ns1)

        ns2 = NumberStyle(name="Integer")
        ns2.addElement(TextProperties(color="#ff0000"))
        ns2.addElement(Text(text="-"))
        ns2.addElement(Number(decimalplaces="0", minintegerdigits="1", grouping="true"))
        ns2.addElement(Map(condition="value()>=0", applystylename="IntegerBlack"))
        doc.styles.addElement(ns2)
            
        #Decimal 2
        ns1 = NumberStyle(name="Decimal2Black", volatile="true")
        ns1.addElement(Number(decimalplaces="2", minintegerdigits="1", grouping="true"))
        doc.styles.addElement(ns1)

        ns2 = NumberStyle(name="Decimal2")
        ns2.addElement(TextProperties(color="#ff0000"))
        ns2.addElement(Text(text="-"))
        ns2.addElement(Number(decimalplaces="2", minintegerdigits="1", grouping="true"))
        ns2.addElement(Map(condition="value()>=0", applystylename="Decimal2Black"))
        doc.styles.addElement(ns2)
            
        #Decimal 2
        ns1 = NumberStyle(name="Decimal6Black", volatile="true")
        ns1.addElement(Number(decimalplaces="6", minintegerdigits="1", grouping="true"))
        doc.styles.addElement(ns1)

        ns2 = NumberStyle(name="Decimal6")
        ns2.addElement(TextProperties(color="#ff0000"))
        ns2.addElement(Text(text="-"))
        ns2.addElement(Number(decimalplaces="6", minintegerdigits="1", grouping="true"))
        ns2.addElement(Map(condition="value()>=0", applystylename="Decimal6Black"))
        doc.styles.addElement(ns2)
        
        #Boolean
        ns1 = BooleanStyle(name="BooleanBlack", volatile="true")
        ns1.addElement(Boolean())
        doc.styles.addElement(ns1)

        ns2 = BooleanStyle(name="Boolean")
        ns2.addElement(TextProperties(color="#ff0000"))
        ns2.addElement(Boolean())
        ns2.addElement(Map(condition="value()==1", applystylename="BooleanBlack"))
        doc.styles.addElement(ns2)
            

    def generate_ods_styles(self, doc, currencymanager):
        self.__generate_ods_common_styles(doc)
        for o in self.arr:
            o.generate_ods_styles(doc, currencymanager)


## Class to write Libreoffice Calc ODS files
class ODS_Write(ODS):
    def __init__(self, filename):
        ODS.__init__(self, filename)

        self.colors=ODSStyleColorManager()
        self.colors.append(ODSStyleColor("Green", "#9bff9b", True))
        self.colors.append(ODSStyleColor("GrayDark", "#888888", True))
        self.colors.append(ODSStyleColor("GrayLight", "#bbbbbb", True))
        self.colors.append(ODSStyleColor("Orange", "#ffcc99", True))
        self.colors.append(ODSStyleColor("Yellow", "#ffff7f", True))
        self.colors.append(ODSStyleColor("White", "#ffffff", True))
        self.colors.append(ODSStyleColor("Blue", "#9b9bff", True))
        self.colors.append(ODSStyleColor("Red", "#ff9b9b", True))
        self.colors.append(ODSStyleColor("Normal", "#ffffff", False))
        
        self.currencies=ODSStyleCurrencyManager()
        self.currencies.append(ODSStyleCurrency("EUR", "€"))
        self.currencies.append(ODSStyleCurrency("USD", "$"))

    ## Generate color styles and currency styles and the save file
    def save(self, filename=None):
        self.currencies.generate_ods_styles(self.doc)
        self.colors.generate_ods_styles(self.doc, self.currencies)
        ODS.save(self, filename)
        

## Guess style from color and object class
## @param color_or_style String with a Color name or a style. If it's a color returns the style corresponding to the object. If it's a style returns the same style
## @return String with the style 
def guess_ods_style(color_or_style, object):
    if color_or_style in ["Green", "GrayDark", "GrayLight", "Orange", "Yellow", "White", "Blue", "Red", "Normal"]:
        if object.__class__.__name__=="str":
            return color_or_style + "Left"
        elif object.__class__.__name__=="int":
            return color_or_style + "Integer"
        elif object.__class__.__name__ in ["Currency", "Money" ]:
            return color_or_style + object.currency
        elif object.__class__.__name__=="Percentage":
            return color_or_style + "Percentage"
        elif object.__class__.__name__ in ("Decimal", "float"):
            return color_or_style +  "Decimal2"
        elif object.__class__.__name__=="datetime":
            return color_or_style + "Datetime"
        elif object.__class__.__name__=="date":
            return color_or_style + "Date"
        elif object.__class__.__name__=="time":
            return color_or_style + "Time"
        elif object.__class__.__name__=="bool":
            return color_or_style + "Boolean"
        else:
            info("guess_ods_style not guessed {}".format( object.__class__))
            return "NormalLeft"
    else:
        return color_or_style

## Gets a ODS file and rewrites it with libreoffice convert-to command
## Can be used to assign data values formulas to file. Or to fix ploblems on specific files.
## @param filename_from canbe an ods or a xlsx file.
## @param filename_to Must be a ods file path
## Returns the name of the recently created file
def create_rewritten_ods(filename_from, filename_to=None):  
    if filename_to is None:
        filename_to=filename_from + ".rewritten.ods"
    convert_command(filename_from, filename_to)
    return filename_to

## Creates a new file with data_only cells (formulas converted to numbers).
## Returns the name of the recently created file
## 1 Transforms to xlsx (temporal_xlsx_rewriten)
## 2 Transforms to data_only xlsx (temporal_xlsx_data_only)
def create_data_only_ods( filename_from, filename_to=None):
    if filename_to is None:
        filename_to=filename_from + ".data_only.ods"
    with TemporaryDirectory(prefix="officegenerator_") as tmp_name:
        from officegenerator.libxlsxgenerator import create_data_only_xlsx
        temporal_xlsx=tmp_name+ sep + path.basename(filename_from+".xlsx")
        create_data_only_xlsx(filename_from, temporal_xlsx)
        convert_command(temporal_xlsx, filename_to)
    return filename_to
        
