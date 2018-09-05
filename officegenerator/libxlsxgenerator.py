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

from officegenerator.commons import columnAdd, makedirs, rowAdd, deprecated
from officegenerator.libodfgenerator import OdfMoney, OdfPercentage
from decimal import Decimal


try:
    t=gettext.translation('officegenerator',pkg_resources.resource_filename("officegenerator","locale"))
    _=t.gettext
except:
    _=str

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
        self.stGreyLight=openpyxl.styles.Color('FFDCDCDC')
        self.stGreyDark=openpyxl.styles.Color('FFC3C3C3')

    ## Freezes panels
    ## @param strcell String For example "A2"
    def freezePanels(self, strcell):
        self.ws_current.freeze_panes=self.ws_current.cell(strcell)
        
    ## Selects a cell
    ## @param strcell String For example "A2"
    def setSelectedCell(self, strcell):
        self.ws_current.sheet_view.pane.topLeftCell=strcell
        self.ws_current.sheet_view.selection=[]
        self.ws_current.sheet_view.selection.append(openpyxl.worksheet.Selection("topLeft", strcell, None, strcell))
        
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
        return self.wb.sheetnames()[id]


    ## It returns a sheet object with the index id
    def get_sheet_by_id(self, id):
        return self.wb[self.wb.sheetnames[id]]

    ## It returns a index integer of the sheet with a given name
    def get_sheet_id(self, name):
        for id, s_name in enumerate(self.wb.sheetnames):
            if s_name==name:
                return id
        return None

    ## After removing it sets current sheet to 0 index
    def remove_sheet_by_id(self, id):
        ws=self.get_sheet_by_id(id)
        self.wb.remove_sheet(ws)
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
    def cell(self,letter,number):
        return self.ws_current[letter+number]

    ## Internat function to set a cell. All properties except border that it's setted in overwrite functions (merged and no merged)
    ## @param cell is a cell object
    def __setValue(self, cell, value, style, decimals, alignment):
        if value.__class__ in (int, ):#Un solo valor
            cell.value=value
            cell.number_format="#.###;[RED]-#.###"
            alignment='right' if alignment==None else alignment
        elif value.__class__ in (float, Decimal):#Un solo valor
            cell.value=value
            alignment='right' if alignment==None else alignment
            cell.number_format="#.##0,00;[RED]-#.##0,00"
        elif value.__class__ in (str, ):#Un solo valor
            cell.value=value
            alignment='left' if alignment==None else alignment
        elif value.__class__ in (datetime.datetime, ):
            cell.value=value
            alignment='left' if alignment==None else alignment
            cell.number_format="YYYY-MM-DD HH:mm"
        elif value.__class__ in (datetime.date, ):
            cell.value=value
            alignment='left' if alignment==None else alignment
            cell.number_format="YYYY-MM-DD"
        elif value.__class__ in (OdfMoney, ):
            cell.value=value.amount
            alignment='right' if alignment==None else alignment
            #zeros=decimals*"0"
            cell.number_format="#.##0,00 {0};[RED]-#.##0,00 {0}".format(value.currency)
            print(cell.number_format)
        elif value.__class__ in (OdfPercentage, ):
            cell.value=value.value
            alignment='right' if alignment==None else alignment
            cell.number_format="#.##0,00 %;[RED]-#.##0,00 %"
        elif value==None:
            return
        else:
            print(value.__class__, "VALUE CLASS NOT FOUND")
        if style!=None:
            cell.fill=openpyxl.styles.PatternFill("solid", fgColor=style)
        bold=False if style==None else True
        cell.font=openpyxl.styles.Font(name='Arial', size=10, bold=bold)
        cell.alignment=openpyxl.styles.Alignment(horizontal=alignment, vertical='center')

    ## Writes a cell
    ## @param alignment String None by default. Can be "right","left","center"
    ## @param style its a openpyxl.styles.Color object. There are several predefined stGreen, stGreyDark, stGreyLight, stOrange, stYellow
    ## @param decimals Integer with the number of decimals. 2 by default
    def overwrite(self, letter, number, result, style=None,  decimals=2, alignment=None):
        if result.__class__== list:#Una lista
            for i,row in enumerate(result):
                if row.__class__ in (list, ):#Una lista de varias columnas
                    for j,column in enumerate(row):
                        cell=self.cell(columnAdd(letter, j), rowAdd(number, i))
                        self.__setValue(cell, result[i][j], style, decimals, alignment)
                        self.__setBorder(cell)
                else:#Una lista de una columna
                    cell=self.cell(letter, rowAdd(number, i))
                    self.__setValue(cell, result[i], style, decimals, alignment)
                    self.__setBorder(cell)
        else:#Un solo valor
            cell=self.cell(letter, number)
            self.__setValue(cell, result, style, decimals, alignment)
            self.__setBorder(cell)


    #Sets border to a cell not merged
    def __setBorder(self, cell):
        cell.border=openpyxl.styles.Border(
            left=openpyxl.styles.Side(border_style='thin'),
            top=openpyxl.styles.Side(border_style='thin'),
            right=openpyxl.styles.Side(border_style='thin'),
            bottom=openpyxl.styles.Side(border_style='thin') 
        )

    ## Sets cell name to use in formulas
    def setCellName(self, range, name):
        self.wb.create_named_range(name, self.ws_current, range)

    ## Set columns width in current sheet
    ## @param arrWidths List with integers representing column width
    def setColumnsWidth(self, arrWidths):
        for i in range(len(arrWidths)):
            self.ws_current.column_dimensions[columnAdd("A", i)].width=arrWidths[i]

    ## Function to merge cells 
    ## @param range String for example: A1:B1
    ## @param style was added to avoid formating errors after merging
    @deprecated
    def mergeCells(self, range, style=None, decimals=None, alignment=None):
        for row in self.ws_current[range]: #Returns a list of cells
            for cell in row:
                self.__setValue_by_cell(cell, cell.value, style, decimals, alignment)
        self.ws_current.merge_cells(range)


    ## Result must be the value only for the first cell
    def overwrite_and_merge(self, cell_range,  result, style=None,  decimals=2, alignment=None):
        self.ws_current.merge_cells(cell_range)
        top = openpyxl.styles.Border(top=openpyxl.styles.Side(border_style='thin'))
        left = openpyxl.styles.Border(left=openpyxl.styles.Side(border_style='thin'))
        right = openpyxl.styles.Border(right=openpyxl.styles.Side(border_style='thin'))
        bottom = openpyxl.styles.Border(bottom=openpyxl.styles.Side(border_style='thin'))

        first_cell = self.ws_current[cell_range.split(":")[0]]
        self.__setValue(first_cell, result, style, decimals, alignment)

        rows = self.ws_current[cell_range]

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
    def setComment(self, strcell, comment):
        self.ws_current[strcell].comment=openpyxl.comments.Comment(comment, "PySGAE")
