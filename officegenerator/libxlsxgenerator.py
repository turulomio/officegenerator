## @namespace officegenerator.libxlsxgenerator
## @brief Este módulo permite la lectura y escritura de ficheros xlsx de Microsoft Excel
##
## You can change current sheet, with createSheet or using setCurrentSheet. After that all commands will use that sheet until you change it again
import datetime
import gettext
import openpyxl
import openpyxl.comments
import openpyxl.cell
import openpyxl.styles
import openpyxl.worksheet
import openpyxl.formatting.rule
import os
import pkg_resources

from officegenerator.commons import columnAdd, makedirs,  Currency,  Percentage,  Coord, Range
from decimal import Decimal


try:
    t=gettext.translation('officegenerator',pkg_resources.resource_filename("officegenerator","locale"))
    _=t.gettext
except:
    _=str


class ColumnWidthXLSX:
    Date=40
    Detetime=60

class OpenPyXL:
    def __init__(self,filename,template=None):
        self.filename=filename
        self.template=template
        if template==None:
            self.wb=openpyxl.Workbook()
        else:
            self.wb=openpyxl.load_workbook(self.template)

        self.ws_current=self.wb.active
        self.setCurrentSheet(self.ws_current.title)

        self.stOrange=openpyxl.styles.Color('FFFFDCA8')
        self.stYellow=openpyxl.styles.Color('FFFFFFC0')
        self.stGreen=openpyxl.styles.Color('FFC0FFC0')
        self.stGrayLight=openpyxl.styles.Color('FFDCDCDC')
        self.stGrayDark=openpyxl.styles.Color('FFC3C3C3')
        self.stWhite=openpyxl.styles.Color('FFFFFFFF')
    
    ## Returns the style name of a givenven color
    ## @param openpyxl.styles.Color
    ## @return string
    def styleName(self, color):
        if color==self.stOrange:
            return "Orange"
        elif color==self.stYellow:
            return "Yellow"
        elif color==self.stGreen:
            return "Green"
        elif color==self.stGrayDark:
            return "Dark gray"
        elif color==self.stGrayLight:
            return "Light gray"
        elif color==self.stWhite:
            return "White"
        elif color==None:
            return "Normal"

    ## Freezes panels
    ## @param strcell String For example "A2"
    def freezePanels(self, coord_string):
        self.ws_current.freeze_panes=self.ws_current[coord_string]

    ## Selects a cell
    ## @param coord_string String For example "A2"
    def setSelectedCell(self, coord_string):
        self.ws_current.views.sheetView[0].selection[0].activeCell=coord_string

    ## Changes name of the current sheet
    def setSheetName(self, name):
        self.ws_current.title=name

    ## Create a sheet at the end, renames it and selects it as current
    def createSheet(self, name):
        self.wb.create_sheet(title=name)
        self.setCurrentSheet(name)

    ## Function that establishes current worksheet. Updates self.ws_current and self.ws_current_id
    ##
    ## id Is a integer beginning with 0
    ## name is the title of the sheet
    ## @param id_or_name Index or Nmae
    def setCurrentSheet(self, id_or_name):
        if id_or_name.__class__==int:
            self.ws_current_id=id_or_name
        else:#name
            self.ws_current_id=self.get_sheet_id(id_or_name)
        self.ws_current=self.get_sheet_by_id(self.ws_current_id)

    def setColorScale(self, range):
        self.ws_current.conditional_formatting.add(range, 
                            openpyxl.formatting.rule.ColorScaleRule(
                                                start_type='percentile', start_value=0, start_color='00FF00',
                                                mid_type='percentile', mid_value=50, mid_color='FFFFFF',
                                                end_type='percentile', end_value=100, end_color='FF0000'
                                                )
                                            )
    ## Returns sheet_name
    def sheet_name(self, id=None):
        if id==None:
            id=self.ws_current_id
        return self.wb.sheetnames[id]


    ## It returns a sheet object with the index id
    def get_sheet_by_id(self, id):
        return self.wb[self.wb.sheetnames[id]]

    ## It returns a index integer of the sheet with a given name
    def get_sheet_id(self, name):
        for id, s_name in enumerate(self.wb.sheetnames):
            if s_name==name:
                return id
        return None

    ## Returns the number of columns with data of the current sheet. Returns the number not the index
    ## @return int
    def max_columns(self):
        return self.ws_current.max_column
        
    ## Returns the number of rows with data of the current sheet. Returns the number not the index
    ## @return int
    def max_rows(self):
        return self.ws_current.max_row
        
        
    ## After removing it sets current sheet to 0 index
    def remove_sheet_by_id(self, id):
        ws=self.get_sheet_by_id(id)
        self.wb.remove(ws)
        self.setCurrentSheet(0)

    def save(self, filename=None):
        if filename==None:
            filename=self.filename
        makedirs(os.path.dirname(filename))
        self.wb.save(filename)

        if os.path.exists(filename)==False:
            print(_("*** ERROR: File wasn't generated ***"))

    ## Returns a cell object in the current sheet
    ## @param letter
    ## @param number
    ## @return sheet
    def cell(self, coord):
        coord=Coord.assertCoord(coord)
        return self.ws_current[coord.string()]

    ## Internal function to set the number format
    ##
    ## This strings are openpyxl string not libreoffice cell string
    ## @param cell is a cell object
    ## @param value Value to add to the cell
    ## @param style Color or None. If None this function it's ignored
    ## @param decimals Number of decimals
    def __setNumberFormat(self, cell, value, style, decimals):     
        if style==None:
            return
        if value.__class__ in (int, ):#Un solo valor
            cell.number_format='#,##0;[RED]-#,##0'
        elif value.__class__ in (float, Decimal):#Un solo valor
            zeros=decimals*"0"
            cell.number_format="#,##0.{0};[RED]-#,##0.{0}".format(zeros)
        elif value.__class__ in (datetime.datetime, ):
            cell.number_format="YYYY-MM-DD HH:mm"
        elif value.__class__ in (datetime.date, ):
            cell.number_format="YYYY-MM-DD"
        elif value.__class__ in (Currency, ):
            cell.number_format='#,##0.00 "{0}";[RED]-#,##0.00 "{0}"'.format(value.symbol())
        elif value.__class__ in (Percentage, ):
            cell.number_format="#.##0,00 %;[RED]-#.##0,00 %"

    ## Internat function to set a cell. All properties except border that it's setted in overwrite functions (merged and no merged)
    ## @param cell is a cell object
    def __setValue(self, cell, value, style, decimals, alignment):     
        if value==None:
            return
        elif value.__class__ in (Currency, ):
            cell.value=value.amount
        elif value.__class__ in (Percentage, ):
            cell.value=value.value
        else:
            cell.value=value


    ## Internal method to set a not merged cell
    ## @param cell is a cell object
    ## @param value Value to add to the cell
    ## @param style Color or None. If None this function it's ignored
    ## @param decimals Number of decimals
    ## @param alignment Cell alignment
    def __setCell(self, coord, value, style=None, decimals=2, alignment=None):
        coord=Coord.assertCoord(coord)
        cell=self.cell(coord.string())
        self.__setValue(cell, value, style, decimals, alignment)
        self.__setBorder(cell, style)
        self.__setAlignment(cell, value, style, alignment)
        self.__setNumberFormat(cell, value, style, decimals)      

        if style!=None:
            cell.fill=openpyxl.styles.PatternFill("solid", fgColor=style)
            bold=False if style==self.stWhite else True
            cell.font=openpyxl.styles.Font(name='Arial', size=10, bold=bold)

    ## Internat function to set cell alignment
    ## @param cell is a cell object
    ## @style Color or None. This method is ignored if style=None
    def __setAlignment(self, cell, value,  style, alignment):  
        if style==None:
            return
        if alignment==None:
            if value.__class__ in (str, datetime.datetime, datetime.date):#Un solo valor
                alignment='left'
            else:
                alignment='right'
        cell.alignment=openpyxl.styles.Alignment(horizontal=alignment, vertical='center')

    ## Writes a cell or a list of cell or a list of list of cells
    ## @param coord Can be a Coord or a string with text coords
    ## @param result Can be a value, a list of values or a list of lists of values
    ## @param style its a openpyxl.styles.Color object. There are several predefined stGreen, stGrayDark, stGrayLight, stOrange, stYellow, stWhite or None. None is used to preserve template cell and the value is the only thing will be changed
    ## @param decimals Integer with the number of decimals. 2 by default
    ## @param alignment String None by default. Can be "right","left","center"
    def overwrite(self, coord, result, style=None,  decimals=2, alignment=None):
        coord=Coord.assertCoord(coord)
        if result.__class__== list:#Una lista
            for i,row in enumerate(result):
                if row.__class__ in (list, ):#Una lista de varias columnas
                    for j,column in enumerate(row):
                        self.__setCell(Coord(coord.string()).addRow(i).addColumn(j), result[i][j], style, decimals, alignment )   
                else:#Una lista de una columna
                    self.__setCell(Coord(coord.string()).addRow(i), result[i], style, decimals, alignment )
        else:#Un solo valor
            self.__setCell(coord, result, style, decimals, alignment )


    ##Sets border to a cell not merged
    ## @param cell is a cell object
    ## @param style Color or None. If None this function it's ignored
    def __setBorder(self, cell, style):
        if style==None:
            return
        cell.border=openpyxl.styles.Border(
            left=openpyxl.styles.Side(border_style='thin'),
            top=openpyxl.styles.Side(border_style='thin'),
            right=openpyxl.styles.Side(border_style='thin'),
            bottom=openpyxl.styles.Side(border_style='thin') 
        )

    ## Sets cell name to use in formulas. Fails if range_string is not fixed. For example: $A$4
    ## @param range_string 
    ## @param name
    def setCellName(self, range_string, name):
        self.wb.create_named_range(name, self.ws_current, range_string)

    ## Set columns width in current sheet
    ## @param arrWidths List with integers representing column width
    def setColumnsWidth(self, arrWidths):
        for i in range(len(arrWidths)):
            self.ws_current.column_dimensions[columnAdd("A", i)].width=arrWidths[i]

    ## Create a merged cell
    ## @param range_string Can be a Range or a range string 
    ## @param value Can be a value. Must be the value only for the first cell
    ## @param style its a openpyxl.styles.Color object. There are several predefined stGreen, stGrayDark, stGrayLight, stOrange, stYellow, stWhite or None. None is used to preserve template cell and the value is the only thing will be changed
    ## @param decimals Integer with the number of decimals. 2 by default
    ## @param alignment String None by default. Can be "right","left","center"
    def overwrite_and_merge(self, range,  value, style=None,  decimals=2, alignment=None):
        range=Range.assertRange(range)
        self.ws_current.merge_cells(range.string())
        top = openpyxl.styles.Border(top=openpyxl.styles.Side(border_style='thin'))
        left = openpyxl.styles.Border(left=openpyxl.styles.Side(border_style='thin'))
        right = openpyxl.styles.Border(right=openpyxl.styles.Side(border_style='thin'))
        bottom = openpyxl.styles.Border(bottom=openpyxl.styles.Side(border_style='thin'))

        self.__setCell(range.start.string(), value, style, decimals, alignment)

        rows = self.ws_current[range.string()]

        for cell in rows[0]:
            cell.border = cell.border + top
        for cell in rows[-1]:
            cell.border = cell.border + bottom

        for row in rows:
            l = row[0]
            r = row[-1]
            l.border = l.border + left
            r.border = r.border + right

    ## Sets a comment
    ## @param strcell String "A1" for example
    def setComment(self, coord_string, comment):
        self.ws_current[coord_string].comment=openpyxl.comments.Comment(comment, "PySGAE")

