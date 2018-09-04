#!/usr/bin/python3
## @namespace officegenerator.libxlsxgenerator
## @brief Este módulo permite la lectura y escritura de ficheros xlsx de Microsoft Excel

## I Create styles with color and a suffix: Decimal2, Integer, Percentage, Euros. For example: GreyInteger
## I Create styles with color and alignment for text. OrangeLeft, OrangeCenter, OrangeRight

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

from officegenerator.commons import columnAdd, makedirs, rowAdd
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

        self.ws_current_id=id
        
        self.stOrange=openpyxl.styles.Color('FFFFDCA8')
        self.stYellow=openpyxl.styles.Color('FFFFFFC0')
        self.stGreen=openpyxl.styles.Color('FFC0FFC0')
        self.stGreyLight=openpyxl.styles.Color('FFDCDCDC')
        self.stGreyDark=openpyxl.styles.Color('FFC3C3C3')

    def freezePanels(self, cell):
        """Cell is string coordinates"""
        ws=self.get_sheet_by_id(self.ws_current_id)
        ws.freeze_panes=ws.cell(cell)
        
    def setSelectedCell(self, cell):
        """Cell is string coordinates
        
        Estaq función fue echa a modo prueba error"""
        ws=self.get_sheet_by_id(self.ws_current_id)
        ws.sheet_view.pane.topLeftCell=cell
        ws.sheet_view.selection=[]
        ws.sheet_view.selection.append(openpyxl.worksheet.Selection("topLeft", cell, None, cell))
        

    def setSheetName(self, name):
        """Changes current id"""
        ws=self.get_sheet_by_id(self.ws_current_id)
        ws.title=name

    def createSheet(self, name):
        """Create a sheet at the end, renames it and selects it as current"""
        self.wb.create_sheet(title=name)
        self.ws_current_id=self.get_sheet_id(name)
        
    def setColorScale(self, range):
        ws=self.get_sheet_by_id(self.ws_current_id)
        ws.conditional_formatting.add(range, 
                            openpyxl.formatting.rule.ColorScaleRule(
                                                start_type='percentile', start_value=0, start_color='00FF00',
                                                mid_type='percentile', mid_value=50, mid_color='FFFFFF',
                                                end_type='percentile', end_value=100, end_color='FF0000'
                                                )   
                                            )

    def sheet_name(self, id=None):
        if id==None:
            id=self.ws_current_id
        return self.wb.sheetnames()[id]

    def get_sheet_by_id(self, id):
        return self.wb[self.wb.sheetnames[id]]

    def get_sheet_id(self, name):
        for id, s_name in enumerate(self.wb.sheetnames):
            if s_name==name:
                return id
        return None

    def remove_sheet_by_id(self, id):
        ws=self.get_sheet_by_id(id)
        self.wb.remove_sheet(ws)
        self.get_sheet_by_id(0)

    def save(self, filename=None):
        if filename==None:
            filename=self.filename
        makedirs(os.path.dirname(filename))
        self.wb.save(filename)
        
        if os.path.exists(filename)==False:
            print( "*** ERROR: El fichero no se ha generado ***")
            
    def overwrite_cell(self):
        print("Hacer para pasar custom cells")

    ## Writes a cell
    ## @param alignment String None by default. Can be "right","left","center"
    ## @param style its a openpyxl.styles.Color object. There are several predefined stGreen, stGreyDark, stGreyLight, stOrange, stYellow
    ## @param decimals Integer with the number of decimals. 2 by default
    def overwrite(self, letter, number, result, ws=None, style=None,  decimals=2, alignment=None):
        def setValue(cell, value, style, decimals, alignment):
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
            cell.border=openpyxl.styles.Border(
                left=openpyxl.styles.Side(border_style='thin'),
                top=openpyxl.styles.Side(border_style='thin'),
                right=openpyxl.styles.Side(border_style='thin'),
                bottom=openpyxl.styles.Side(border_style='thin') 
            )
            cell.alignment=openpyxl.styles.Alignment(horizontal=alignment, vertical='center')
        ################################################
        if ws==None:
            ws=self.get_sheet_by_id(self.ws_current_id)
        if result.__class__== list:#Una lista
            for i,row in enumerate(result):
                if row.__class__ in (list, ):#Una lista de varias columnas
                    for j,column in enumerate(row):
                        c=ws[columnAdd(letter, j)+rowAdd(number,i)]
                        setValue(c, result[i][j], style, decimals, alignment)
                else:#Una lista de una columna
                    c=ws[letter+rowAdd(number,i)]
                    setValue(c, result[i], style, decimals, alignment)
        else:#Un solo valor
            c=ws[letter+number]
            setValue(c, result, style, decimals, alignment)

    ## Sets cell name to use in formulas
    def setCellName(self, range, name):
        ws=self.get_sheet_by_id(self.ws_current_id)
        self.wb.create_named_range(name, ws, range)

    ## Sets the corrent sheet 
    ## @param id Integer with the index of the sheet
    def setCurrentSheet(self, id):
        self.ws_current_id=id
        
    ## Set columns width in current sheet
    ## @param arrWidths List with integers representing column width
    def setColumnsWidth(self, arrWidths):
        ws=self.get_sheet_by_id(self.ws_current_id)
        for i in range(len(arrWidths)):
            ws.column_dimensions[columnAdd("A", i)].width=arrWidths[i]
   
    def mergeCells(self, range, style=None):
        """range='A1:B1
        
        style was added to avoid formating errors after merging
        
        """
        ws=self.get_sheet_by_id(self.ws_current_id)
        if style!=None:
            for row in ws.iter_rows(range):
                for c in row:
                    c.style=style
        ws.merge_cells(range)
        
        
    def setComment(self, cell, comment):
        """Cell is string coordinates"""
        ws=self.get_sheet_by_id(self.ws_current_id)
        ws[cell].comment=openpyxl.comments.Comment(comment, "PySGAE")
