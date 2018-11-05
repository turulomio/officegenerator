## @namespace officegenerator.libodfgenerator
## @brief Package that allows to read and write Libreoffice ods and odt files
## This file is from the Xulpymoney project. If you want to change it. Ask for project administrator

## odf Element and Text inherits from Node and inherits from xml.Dom.Node https://docs.python.org/3.5/library/xml.dom.html
## Text can't have children

import datetime
import gettext
import logging
import os
import pkg_resources
from decimal import Decimal
from odf.opendocument import OpenDocumentSpreadsheet,  OpenDocumentText,  load
from odf.style import Footer, FooterStyle, HeaderFooterProperties, Style, TextProperties, TableColumnProperties, Map,  TableProperties,  TableCellProperties, PageLayout, PageLayoutProperties, ParagraphProperties,  ListLevelProperties,  MasterPage
from odf.number import  CurrencyStyle, CurrencySymbol,  Number, NumberStyle, Text,  PercentageStyle,  DateStyle, Year, Month, Day, Hours, Minutes, Seconds
from odf.text import P,  H,  Span, ListStyle,  ListLevelStyleBullet,  List,  ListItem, ListLevelStyleNumber,  OutlineLevelStyle,  OutlineStyle,  PageNumber,  PageCount
from odf.table import Table, TableColumn, TableRow, TableCell,  TableHeaderRows
from odf.draw import Frame, Image
from odf.dc import Creator, Description, Title, Date, Subject
from odf.meta import InitialCreator
import odf.element

from odf.config import ConfigItem, ConfigItemMapEntry, ConfigItemMapIndexed, ConfigItemMapNamed,  ConfigItemSet
from odf.office import Annotation

from officegenerator.commons import makedirs,  number2column,  number2row,  Coord, Range,  Percentage, Currency,  __version__

try:
    t=gettext.translation('officegenerator',pkg_resources.resource_filename("officegenerator","locale"))
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
        
        
    def rowNumber(self, sheet_element):
        """
            Devuelve el numero de filas de un determinado sheet_element
        """
        return len(sheet_element.getElementsByType(TableRow))-1
        
    def columnNumber(self, sheet_element):
        """
            Devuelve el numero de filas de un determinado sheet_element
        """
        return len(sheet_element.getElementsByType(TableColumn))
        
    ## Returns the cell value
    def getCellValue(self, sheet_element, coord):
        coord=Coord.assertCoord(coord)
        row=sheet_element.getElementsByType(TableRow)[coord.numberIndex()]
        cell=row.getElementsByType(TableCell)[coord.letterIndex()]
        r=None
        
        if cell.getAttribute('valuetype')=='string':
            r=cell.getAttribute('value')
        if cell.getAttribute('valuetype')=='float':
            r=Decimal(cell.getAttribute('value'))
        if cell.getAttribute('valuetype')=='percentage':
            r=Percentage(Decimal(cell.getAttribute('value')), Decimal(1))
        if cell.getAttribute('formula')!=None:
            r=str(cell.getAttribute('formula'))[3:]
        if cell.getAttribute('valuetype')=='currency':
            r=Currency(Decimal(cell.getAttribute('value')), cell.getAttribute('currency'))
        if cell.getAttribute('valuetype')=='date':
            datevalue=cell.getAttribute('datevalue')
            if len(datevalue)<=10:
                r=datetime.datetime.strptime(datevalue, "%Y-%m-%d").date()
            else:
                r=datetime.datetime.strptime(datevalue, "%Y-%m-%dT%H:%M:%S")
        return r

    ## Returns an odfcell object
    def getCell(self, sheet_element,  coord):
        coord=Coord.assertCoord(coord)
        row=sheet_element.getElementsByType(TableRow)[coord.numberIndex()]
        cell=row.getElementsByType(TableCell)[coord.letterIndex()]
        object=self.getCellValue(sheet_element, coord)
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
        
    def setCell(self, sheet_element,  coord, cell):
        """
            Updates a cell
            insertBefore(newchild, refchild) – Inserts the node newchild before the existing child node refchild.
appendChild(newchild) – Adds the node newchild to the end of the list of children.
removeChild(oldchild) – Re

        ESTA FUNCION SE USA PARA SUSTITUIR EN UNA PLANTILLA
        NO SE PUEDEN AÑADIR MAS CELDAS O FILAS
        PARA ESO USAR ODS_Write DE MOMENTO
        """
        coord=Coord.assertCoord(coord)
        row=sheet_element.getElementsByType(TableRow)[coord.numberIndex()]
        oldcell=row.getElementsByType(TableCell)[coord.letterIndex()]
        row.insertBefore(cell.generate(), oldcell)
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
        
                
    def setMetadata(self, title,  subject, creator):
        for e in self.doc.meta.childNodes:
            self.doc.meta.removeChild(e)
        self.doc.meta.addElement(Description(text=_("This document has been generated with OfficeGenerator v{}".format(__version__))))
        self.doc.meta.addElement(Title(text=title))
        self.doc.meta.addElement(Subject(text=subject))
        self.doc.meta.addElement(Creator(text=creator))
        self.doc.meta.addElement(InitialCreator(text=creator))

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
        self.cursor=None
        self.cursorParent=self.doc.text

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

    ## Search for a tag type elementes  an returns element and its index. With it, inserts new with replaced text and removes old
    ##
    ## Cursor doesn't change because we replace Text objects in Element Text
    ## @param tag String to search
    ## @param replace String to replace. Can't be None
    def search_and_replace(self, tag, replace, type=P):
        e,  textindex=self.search(tag, type) #Places cursor to element
        if e==None:
            print(_("Tag {} hasn't been found"))
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
            e.removeChild(to_remove)

    ## Search for a tag in doc an replaces its elemente with the parameter element
    ##
    ## @param tag String to search
    ## @param replace ELement. OdfPy element
    def search_and_replace_element(self, tag, newelement, type=P):
        e,  textindex=self.search(tag, type) #Places cursor to element
        if e==None:
            print(_("Tag {} hasn't been found"))
            return


        if newelement==None:#Remove paragraph
            print("New element can't be None")
            return

        self.cursorParent.insertBefore(newelement,e)
        self.cursorParent.removeChild(e)
        self.__setCursor(newelement)


    ## Searchs for the item with a tag. Perhaps is its paren where I'll have to append. Only finds the first one
    ## Returns the element p and the position in its text children
    def search(self, tag, type=P):
        for e in self.doc.getElementsByType(type):
            if str(e).find(tag)!=-1:
                self.__setCursor(e)
                for index, child in enumerate(e.childNodes):
                    #print(index, child)
                    if str(child).find(tag)!=-1:
                        #print("SEARCH RETURN",  e, index)
                        return e, index
        print ("tag {} not found".format(tag))
        return None, None


    ## Converts saved odt to pdf. It will have the same file name but with .pdf extension
    def convert_to_pdf(self):
        os.system("lowriter --headless --convert-to pdf '{}'".format(self.filename))

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
        makedirs(os.path.dirname(self.filename))
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
    ## @param header List with all header strings
    ## @param data Multidimension List with all data objects. Can be str, Decimal, int, datetime, date, Currency, Percentage
    ## @param sizes Integer list with sizes in cm
    ## @param fontsize Integer in pt
    ## @param name str or None. Sets the object name. Appears in LibreOffice navigator. If none table will be named to "Table.Sequence"
    ## @param after True: insert after self.cursor element. False: insert before self.cursor element. None: Just return element
    def table(self, header, data, sizes, fontsize, name=None, after=True):
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
            s=Span(text=o)
            if o.__class__ in (str, datetime.datetime, datetime.date ):
                p = P(stylename="Table.Contents.Font{}".format(fontsize))
                s=Span(text=str(o))
            elif o.__class__ in (Currency,  Percentage):
                if o.isLTZero():
                    p = P(stylename="Table.ContentsRight.FontRed{}".format(fontsize))
                s=Span(text=o.string())
            elif o.__class__ in (int, Decimal,  float):
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
        for i, head in enumerate(header):
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
        template=pkg_resources.resource_filename("officegenerator","templates/odt/standard.odt")
        ODT.__init__(self, filename, template, language, country)

    ## Creates a text header
    ## @param text String with the header string
    ## @Level Integer Level of the header
    def header(self, text, level, after=True):
        h=H(outlinelevel=level, stylename="Heading_20_{}".format(level), text=text)
        return self.insertInCursor(h, after)

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
        if self.object.__class__==Currency:
            odfcell = TableCell(valuetype="currency", currency=self.object.currency, value=self.object.amount, stylename=self.style)
        elif self.object.__class__==Percentage:
            odfcell = TableCell(valuetype="percentage", value=self.object.value, stylename=self.style)
        elif self.object.__class__==datetime.datetime:
            odfcell = TableCell(valuetype="date", datevalue=self.object.strftime("%Y-%m-%dT%H:%M:%S"), stylename=self.style)
        elif self.object.__class__==datetime.date:
            odfcell = TableCell(valuetype="date", datevalue=str(self.object), stylename=self.style)
        elif self.object.__class__ in (Decimal, float,  int):
            odfcell= TableCell(valuetype="float", value=self.object,  stylename=self.style)
        else:#strings
            if len(self.object)>0:
                if self.object[:1]=="=":#Formula
                    odfcell = TableCell(formula="of:"+self.object,  stylename=self.style)
                else:
                    odfcell = TableCell(valuetype="string", value=self.object,  stylename=self.style)
                    odfcell.addElement(P(text = self.object))
            else:#Cadena
                odfcell = TableCell(valuetype="string", value=self.object,  stylename=self.style)
                odfcell.addElement(P(text = self.object))
        if self.spannedRows!=1 or self.spannedColumns!=1:
            odfcell.setAttribute("numberrowsspanned", str(self.spannedRows))
            odfcell.setAttribute("numbercolumnsspanned", str(self.spannedColumns))
        if self.comment!=None:
            a=Annotation(textstylename="Right")
            d=Date()
            d.addText(datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"))
            a.addElement(d)
            a.addElement(P(stylename="Right", text=self.comment))
            odfcell.addElement(a)
        return odfcell

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
        self.setCursorPosition("A1")#Default values
        self.setSplitPosition("A1")


    def setSplitPosition(self, coord):
        """
                split/freeze vertical (0|1|2) - 1 = split ; 2 = freeze
    split/freeze horizontal (0|1|2) - 1 = split ; 2 = freeze
    vertical position = in cell if fixed, in screen unit if frozen
    horizontal position = in cell if fixed, in screen unit if frozen
    active zone in the splitted|frozen sheet (0..3 from let to right, top
to bottom)


#   COMPROBADO CON ODF2XML
B1: 
              <config:config-item config:name="HorizontalSplitMode" config:type="short">2</config:config-item>
              <config:config-item config:name="VerticalSplitMode" config:type="short">0</config:config-item>
              <config:config-item config:name="HorizontalSplitPosition" config:type="int">1</config:config-item>
              <config:config-item config:name="VerticalSplitPosition" config:type="int">0</config:config-item>
              <config:config-item config:name="ActiveSplitRange" config:type="short">3</config:config-item>
              <config:config-item config:name="PositionLeft" config:type="int">0</config:config-item>
              <config:config-item config:name="PositionRight" config:type="int">1</config:config-item>
              <config:config-item config:name="PositionTop" config:type="int">0</config:config-item>
              <config:config-item config:name="PositionBottom" config:type="int">0</config:config-item>

"""
        def setActiveSplitRange():
            """
                Creo que es la posición tras los ejes.
            """
            if (self.horizontalSplitPosition!="0" and self.verticalSplitPosition=="0"):
                return "3"
            if (self.horizontalSplitPosition=="0" and self.verticalSplitPosition!="0"):
                return "2"
            if self.horizontalSplitPosition!="0" and self.verticalSplitPosition!="0":
                return "3"
            return "2"

        coord=Coord.assertCoord(coord)
        self.horizontalSplitPosition=str(coord.letterIndex())
        self.verticalSplitPosition=str(coord.numberIndex())
        self.horizontalSplitMode="0" if self.horizontalSplitPosition=="0" else "2"
        self.verticalSplitMode="0" if self.verticalSplitPosition=="0" else "2"
        self.activeSplitRange=setActiveSplitRange()
        self.positionTop="0"
        self.positionBottom="0" if self.verticalSplitPosition=="0" else str(self.verticalSplitPosition)
        self.positionLeft="0"
        self.positionRight="0" if self.horizontalSplitPosition=="0" else str(self.horizontalSplitPosition)

    def setCursorPosition(self, coord):
        """
            Sets the cursor in a Sheet
        """
        coord=Coord.assertCoord(coord)
        self.cursorPositionX=coord.letterIndex()
        self.cursorPositionY=coord.numberIndex()

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
            if c.coord.string()==coord.string():
                return c
        return None


    ## Adds a cell to the sheet using its coord, an object and a color or a style
    ## @param coord Coord where the cell is going to be created
    ## @param result Object to add to the Cell. Can be int, str, float, datetime.datetime, datetime.date, Currency, Percentage, Decimal, None (will be converted to " - ")
    ## @param color_or_style String with a color: Normal, White, Yellow, Orange, Blue, Red, GrayLight, GrayDark. Or a style WhiteInteger, YellowLeft, OrangeCenter, OrangeEUR, RedPercentage...
    def add(self, coord, result, color_or_style="Normal"):
        coord=Coord.assertCoord(coord)

        if result.__class__ in (list,):#Una lista
            for i,row in enumerate(result):
                if row.__class__ in (list, ):#Una lista de varias columnas
                    for j, column in enumerate(row):
                        style=guess_ods_style(color_or_style, result[i][j])
                        self.addCell(OdfCell(Coord(coord.string()).addColumn(j).addRow(i), result[i][j], style))
                else: #Any value not list if row.__class__ in (int, str, float, datetime.datetime,  datetime.date, Currency, Percentage,  Decimal):#Una lista de una columna
                    style=guess_ods_style(color_or_style, result[i])
                    self.addCell(OdfCell(Coord(coord.string()).addRow(i), result[i], style))
        else: #Any value not list#result.__class__ in (str, int, float, datetime.datetime, datetime.date,  Currency, Percentage, Decimal):#Un solo valor
            style=guess_ods_style(color_or_style, result)  
            self.addCell(OdfCell(coord, result, style))

    ## Adds a cell to self.arr with merge, content and style information
    ## @param range Range
    def addMerged(self, range, result, style):
        range=Range.assertRange(range)
        self.add(range.start, result, style)      
        c=self.getCell(range.start)
        c.setSpanning(range.numColumns(), range.numRows())

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
                if cell!=None:
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

        makedirs(os.path.dirname(filename))
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

        date = Style(name=self.name+"Date", datastylename="Date",parentstylename=self.name+"Left", family="table-cell")
        doc.styles.addElement(date)

        integer = Style(name=self.name+"Integer", family="table-cell",  datastylename="Integer",parentstylename=self.name+"Right")
        doc.styles.addElement(integer)

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
        if object.__class__==str:
            return color_or_style + "Left"
        elif object.__class__==int:
            return color_or_style + "Integer"
        elif object.__class__==Currency:
            return color_or_style + object.currency
        elif object.__class__==Percentage:
            return color_or_style + "Percentage"
        elif object.__class__ in (Decimal, float):
            return color_or_style +  "Decimal2"
        elif object.__class__==datetime.datetime:
            return color_or_style + "Datetime"
        elif object.__class__==datetime.date:
            return color_or_style + "Date"
        else:
            logging.info("guess_ods_style not guessed",  object.__class__)
            return "NormalLeft"
    else:
        return color_or_style
