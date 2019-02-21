## @namespace officegenerator.demo
## @brief Generate ODF example files

import argparse
import datetime
import gettext
import os
import pkg_resources
from decimal import Decimal
from officegenerator.commons import __version__
from officegenerator.libodfgenerator import ODS_Read, ODS_Write, ODT_Manual_Styles, ODT_Standard,  OdfCell, ColumnWidthODS
from officegenerator.libxlsxgenerator import OpenPyXL
from officegenerator.commons import argparse_epilog, Coord, Percentage,  Currency
from odf.text import P
import openpyxl.styles

try:
    t=gettext.translation('officegenerator',pkg_resources.resource_filename("officegenerator","locale"))
    _=t.gettext
except:
    _=str

## If arguments is None, launches with sys.argc parameters. Entry point is toomanyfiles:main
## You can call with main(['--pretend']). It's equivalento to os.system('program --pretend')
## @param arguments is an array with parser arguments. For example: ['--argument','9']. 
def main(arguments=None):
    parser=argparse.ArgumentParser(prog='officegenerator', description=_('Create example files using officegenerator module'), epilog=argparse_epilog(), formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version=__version__)
    group= parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--create', help="Create demo files", action="store_true",default=False)
    group.add_argument('--remove', help="Remove demo files", action="store_true", default=False)
    args=parser.parse_args(arguments)

    if args.remove==True:
        os.remove("officegenerator.ods")
        os.remove("officegenerator.odt")
        os.remove("officegenerator_manual_styles.odt")
        os.remove("officegenerator_readed.ods")
        os.remove("officegenerator.xlsx")
        os.remove("officegenerator_readed.xlsx")

    if args.create==True:
        print(_("Generating example files"))
        demo_ods()
        print("  * " + _("ODS Generated"))

        demo_ods_readed()
        print("  * " + _("ODS Readed and regenerated"))

        demo_odt_standard()
        print("  * " + _("ODT Generated"))

        demo_odt_manual_styles()
        print("  * " + _("ODT Generated from Manual Styles"))

        demo_xlsx()
        print("  * " + _("XLSX Generated"))

        demo_xlsx_readed()
        print("  * " + _("XLSX Readed and regenerated"))


def demo_ods_readed():
    doc=ODS_Read("officegenerator.ods")
    s1=doc.getSheetElementByIndex(0)

    #Sustituye celda
    odfcell=doc.getCell(s1, "B6")
    odfcell.object=1789.12
    odfcell.setComment(_("This cell has been readed and modified"))
    doc.setCell(s1, "B6", odfcell)

    #Added cell
    odfcell=doc.getCell(s1, "B10")
    odfcell.object=_("Created cell")
    odfcell.setComment(_("This cell has been readed and modified"))
    doc.setCell(s1, "B10", odfcell )

    doc.save("officegenerator_readed.ods")

def demo_ods():
    doc=ODS_Write("officegenerator.ods")
    doc.setMetadata("OfficeGenerator example",  "This class documentation", "Mariano Muñoz")
    s1=doc.createSheet("Example")
    s1.add("A1", [["Title", "Value"]], "OrangeCenter")
    s1.add("A2", "Percentage", "YellowLeft")
    s1.add("A4",  "Suma", "WhiteCenter")
    s1.add("B2",  Percentage(12, 56), "WhitePercentage")
    s1.add("B3",  Percentage(12, 21), "WhitePercentage")
    s1.add("B4",  "=sum(B2:B3)","WhitePercentage" )
    s1.add("B6",  100.26, "WhiteDecimal6")
    s1.add("B7",  101, "WhiteInteger")
    s1.setCursorPosition("A3")
    s1.setSplitPosition("A2")

    #Manual cell
    cell=OdfCell("B10", "Celda con OdfCell", "YellowCenter")
    cell.setComment("Comentario")
    cell.setSpanning(2, 2)
    s1.addCell(cell)
    
    #Better way
    s1.addMerged("E10:F11", "Celda con Merged", "GrayDarkCenter")
    s1.setComment("E10", _("This is a comment"))

    #Using lists in arr
    s1.add("A13", [["Una fila"]*3], "Orange")
    s1.add("A15", [[12.3]*3, [12.3]*3])


    s4=doc.createSheet("Splitting")
    for letter in "ABCDEFGHIJ":
        for number in range(1, 11):
            s4.add(letter + str(number), letter+str(number), "YellowLeft")
    s4.setCursorPosition("C3")
    s4.setSplitPosition("C3")


    s6=doc.createSheet("Format number")
    s6.setColumnsWidth([ColumnWidthODS.L, ColumnWidthODS.Datetime, ColumnWidthODS.Date, ColumnWidthODS.L, ColumnWidthODS.L, ColumnWidthODS.L, ColumnWidthODS.L, ColumnWidthODS.XL, ColumnWidthODS.XXL])

    s6.add("A1", _("Style name"), "OrangeCenter")
    s6.add("B1", _("Date and time"), "OrangeCenter")
    s6.add("C1", _("Date"), "OrangeCenter")
    s6.add("D1", _("Integer"), "OrangeCenter")
    s6.add("E1", _("Euros"), "OrangeCenter")
    s6.add("F1", _("Dollars"), "OrangeCenter")
    s6.add("G1", _("Percentage"), "OrangeCenter")
    s6.add("H1", _("Number with 2 decimals"), "OrangeCenter")
    s6.add("I1", _("Number with 6 decimals"), "OrangeCenter")
    for row, color in enumerate(doc.colors.arr):
        s6.add(Coord("A2").addRow(row), color.name, color.name + "Left")
        s6.add(Coord("B2").addRow(row), datetime.datetime.now(),  color.name +"Datetime")
        s6.add(Coord("C2").addRow(row), datetime.date.today(), color.name + "Date")
        s6.add(Coord("D2").addRow(row), pow(-1, row)*-10000000, color.name+ "Integer")
        s6.add(Coord("E2").addRow(row), Currency(pow(-1, row)*12.56, "EUR"), color.name + "EUR")
        s6.add(Coord("F2").addRow(row), Currency(pow(-1, row)*12345.56, "USD"), color.name + "USD")
        s6.add(Coord("G2").addRow(row), Percentage(pow(-1, row)*1, 3), color.name+"Percentage")
        s6.add(Coord("H2").addRow(row), pow(-1, row)*123456789.121212, color.name+"Decimal2")
        s6.add(Coord("I2").addRow(row), pow(-1, row)*-12.121212, color.name+"Decimal6")

    s6.setComment("B2", _("This is a comment"))
    
    #Merge cells
    s6.addMerged("B13:F14", _("This cell is going to be merged with B13 a F14"), "GreenCenter")
    s6.addMerged("B18:G18", _("This cell is going to be merged and aligned desde B18 a G18"), "YellowRight")
    s6.setCursorPosition("B11")
    s6.setSplitPosition("A11")

    doc.setActiveSheet(s6)
    doc.save()
   
def demo_odt_commands(doc):
    doc.setMetadata("OfficeGenerator title",  "OfficeGenerator subject", "Turulomio")
    doc.title(_("Manual of officegenerator"))
    doc.subtitle(_("Version {}".format(__version__)))

    doc.header(_("ODT"), 1)
    doc.simpleParagraph(_("ODT files can be quickly generated with OfficeGenerator.") + " " + 
                                       _("It create predefined styles that allows to create nice documents without worry about styles."))

    doc.header(_("OfficeGenerator predefined paragraph styles"), 2)
    doc.simpleParagraph(_("OfficeGenerator has headers and titles as you can see in the document structure.") + " " + 
                                       _("Morever, it has the following predefined styles:"))
    doc.simpleParagraph(_("This is the 'Standard' style"))
    doc.simpleParagraph(_("This is the 'StandardCenter' style"), style='StandardCenter')
    doc.simpleParagraph(_("This is the 'StandardRight' style"), style='StandardRight')
    doc.simpleParagraph(_("This is the 'Illustration' style"), style='Illustration')
    doc.simpleParagraph(_("This is the 'Bold18Center' style"), style='Bold18Center')
    doc.simpleParagraph(_("This is the 'Bold16Center' style"), style='Bold16Center')
    doc.simpleParagraph(_("This is the 'Bold14Center' style"), style='Bold14Center')
    doc.simpleParagraph(_("This is the 'Bold12Center' style"), style='Bold12Center')
    doc.simpleParagraph(_("This is the 'Bold12Underline' style"), style='Bold12Underline')
    doc.pageBreak()

    doc.header(_("Tables"), 2)
    doc.simpleParagraph(_("We can create tables too, for example with size 11pt:"))
    doc.table(  [_("Concept"), _("Value")], 
                        [
                            [_("Text"), _("This is a text")], 
                            [_("Datetime"), datetime.datetime.now()], 
                            [_("Date"), datetime.date.today()], 
                            [_("Decimal"), Decimal("12.121")], 
                            [_("Currency"), Currency(12.12, "EUR")], 
                            [_("Percentage"), Percentage(1, 3)], 
                        ],
                        [3, 4],
                        11,
                        name=_("First"),
                    )
    doc.simpleParagraph(_("Tables with the size 10pt:"))
    doc.table(  [_("Concept"), _("Value")], 
                        [
                            [_("Text"), _("This is a text")], 
                            [_("Decimal"), Decimal("-12.121")], 
                            [_("Currency"), Currency(-12.12, "EUR")], 
                            [_("Percentage"), Percentage(-1, 3)], 
                        ], 
                        [4, 5], 
                        10
                    )
    doc.simpleParagraph(_("Tables with 8pt size:"))
    doc.table(  [_("Concept"), _("Value")], 
                        [
                            [_("Text"), _("This is a text")], 
                        ], 
                        [4,12], 
                        8,
                        name=_("Third")
                    )
    doc.pageBreak()

    doc.header(_("Lists and numbered lists"), 2) 
    
    doc.simpleParagraph(_("Simple list"))
    doc.list(   [   ["Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. ", ], 
                        ["Adios", ], 
                        ["Bienvenido", ]
                    ],  list_style="List_20_1", paragraph_style="Standard")       
    doc.simpleParagraph(_("Multilevel list with bullets"))
    doc.list(   [   ["1", ["1.1", "1.2"]], 
                        ["2"], 
                        ["3",  ["3.1", ["3.1.1", "3.1.2"]]]
                    ],  list_style="List_20_1",  paragraph_style="Standard")   
    doc.simpleParagraph(_("Multilevel list with Numbers"))
    doc.list(   [   ["1", ["1.1", "1.2"]], 
                        ["2"], 
                        ["3",  ["3.1", ["3.1.1", "3.1.2"]]]
                    ],  list_style="Numbering_20_123", paragraph_style="Text_20_body")   
    doc.pageBreak()

    doc.header(_("Images"), 2)
    pngfile = pkg_resources.resource_filename(__name__, 'images/crown.png')#Gets package filename
    doc.addImage(pngfile,"images/crown.png")#Add image to the doc, adding a key "images/crown.png"
    p = P(stylename="Standard")
    p.addText("Este es un ejemplo de imagen as char: ")
    p.addElement(doc.image("images/crown.png", "3cm", "3cm"))
    p.addText(". Ahora sigo escribiendo sin problemas.")
    doc.insertInCursor(p, after=True)

    doc.simpleParagraph(_("As you can see, I can reuse it one hundred times. File size will not be increased because I used reference names."))
    p=P(stylename="Standard")
    for i in range(100):
        p.addElement(doc.image("images/crown.png", 0.2, 0.2, name="Crown.{}".format(i)))
    doc.insertInCursor(p, after=True)

    doc.simpleParagraph(_("The next paragraph is generated with the illustration method"))
    doc.illustration(["images/crown.png"]*5, 2.5,2.5, "IllustrationCrowns")

    doc.pageBreak(horizontal=True)
    doc.header(_("Horizontal page"), 2)
    doc.simpleParagraph(_("By default OfficeGenerator uses a predefined A4 page.") + " " +  _("Morever, you can set a predefined A4 horizontal page, as you can see."))
    doc.pageBreak()

    doc.header("ODS Writing", 1)
    doc.simpleParagraph("This library create several default styles for writing ODS files. You can see examples in officegenerator.ods.")
    doc.pageBreak()

    doc.header("ODS Reading", 1)
    doc.pageBreak()

    doc.header("XLSX", 1)

    
def demo_odt_standard():
    doc=ODT_Standard("officegenerator.odt")
    demo_odt_commands(doc)
    doc.save()
    
def demo_odt_manual_styles():
    doc=ODT_Manual_Styles("officegenerator_manual_styles.odt")
    demo_odt_commands(doc)
    doc.save()

def demo_xlsx():
    xlsx=OpenPyXL("officegenerator.xlsx")
    xlsx.setCurrentSheet(0)

    xlsx.setSheetName(_("Styles"))
    xlsx.setColumnsWidth([20, 20, 20, 20, 20, 20, 20, 20])
    
    xlsx.overwrite("A1", _("Style name"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("B1", _("Date and time"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("C1", _("Date"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("D1", _("Integer"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("E1", _("Euros"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("F1", _("Percentage"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("G1", _("Number with 2 decimals"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("H1", _("Number with 6 decimals"), style=xlsx.stOrange,  alignment="center")
    for row, style in enumerate([xlsx.stOrange, xlsx.stGreen, xlsx.stGrayLight, xlsx.stYellow, xlsx.stGrayDark, xlsx.stWhite, None]):
        xlsx.overwrite(Coord("A2").addRow(row), xlsx.styleName(style), style=style)
        xlsx.overwrite(Coord("B2").addRow(row), datetime.datetime.now(), style=style)
        xlsx.overwrite(Coord("C2").addRow(row), datetime.date.today(), style=style)
        xlsx.overwrite(Coord("D2").addRow(row), pow(-1, row)*-10000000, style=style)
        xlsx.overwrite(Coord("E2").addRow(row), Currency(pow(-1, row)*12.56, "EUR"), style=style)
        xlsx.overwrite(Coord("F2").addRow(row), Percentage(1, 3), style=style,  decimals=row+1)
        xlsx.overwrite(Coord("G2").addRow(row), pow(-1, row)*12.121212, style=style, decimals=2)
        xlsx.overwrite(Coord("H2").addRow(row), pow(-1, row)*-12.121212, style=style, decimals=6)
    xlsx.setComment("B2", _("This is a comment"))
    
    ##To write a custom cell
    cell=xlsx.wb.active['B12']
    cell.font=openpyxl.styles.Font(name='Arial', size=16, bold=True, color=openpyxl.styles.colors.RED)
    cell.value=_("This is a custom cell")
    #Merge cells
    xlsx.overwrite_and_merge("A13:C14", _("This cell is going to be merged with B13 and C13"),style=xlsx.stOrange)
    xlsx.overwrite_and_merge("A18:G18", _("This cell is going to be merged and aligned"),style=xlsx.stGrayDark, alignment="right")

    xlsx.setSelectedCell("B11")
    xlsx.freezePanels("A9")

    xlsx.overwrite("A20",  [["Una fila"]*3], style=xlsx.stGrayDark)
    xlsx.overwrite_and_merge("E13:G13", _("This sheet max rows are {} and max columns {}").format(xlsx.max_rows(), xlsx.max_columns()), style=xlsx.stYellow,  alignment="center")

    #Named cells
    xlsx.overwrite_and_merge("A23:B23", _("Cell B23 has a name 'Amount"), style=xlsx.stWhite)
    xlsx.overwrite("C23", 5, style=xlsx.stWhite)
    xlsx.setCellName("$C$23", "Amount")

    xlsx.overwrite_and_merge("A24:B24", _("Cell B24 has a name 'Price"), style=xlsx.stWhite)
    xlsx.overwrite("C24", Currency(10,'EUR'), style=xlsx.stWhite)
    xlsx.setCellName("$C$24", "Price")

    xlsx.overwrite_and_merge("A25:B25", _("Cell B25 has a product with names"), style=xlsx.stWhite)
    xlsx.overwrite("C25", "=Amount*Price", style=xlsx.stWhite, alignment='right')



    #To text split and cur position
    xlsx.createSheet("Splitting")
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        for number in range(1, 200):
            xlsx.overwrite(letter + str(number), letter+str(number), style=xlsx.stYellow)
    xlsx.setSelectedCell("Z200")
    xlsx.freezePanels("C3")


    xlsx.save()

def demo_xlsx_readed():
    xlsx=OpenPyXL("officegenerator_readed.xlsx", "officegenerator.xlsx")
    xlsx.setCurrentSheet(0)
    
    xlsx.overwrite("A2", _("Orange"))
    xlsx.overwrite("A5", _("Yellow"))
    xlsx.overwrite("A7", _("White"), style=xlsx.stWhite, alignment="center")
    xlsx.overwrite("D4", 1223)
    #Merge cells
    xlsx.overwrite_and_merge("A20:C21", _("This cell is going to be merged with B13 and C13"),style=xlsx.stOrange)
    xlsx.overwrite_and_merge("A23:G23", _("This cell is going to be merged and aligned"),style=xlsx.stGrayDark, alignment="right")

    xlsx.save()

if __name__ == "__main__":
    main()
