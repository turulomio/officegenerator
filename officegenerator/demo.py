## @namespace officegenerator.demo
## @brief Generate ODF example files

import argparse
import gettext
import pkg_resources
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import timedelta, datetime, date
from decimal import Decimal
from multiprocessing import cpu_count

from officegenerator.commons import __version__, addDebugSystem
from officegenerator.libodfgenerator import ODS_Read, ODS_Write, ODT_Manual_Styles, ODT_Standard,  OdfCell, ColumnWidthODS, ODT, create_rewritten_ods, create_data_only_ods
from officegenerator.libxlsxgenerator import XLSX_Write, XLSX_Read, create_data_only_xlsx, create_rewritten_xlsx
from officegenerator.standard_sheets import Model, Model_Auto
from officegenerator.commons import argparse_epilog, Coord, Range
from officegenerator.objects.currency import Currency
from officegenerator.objects.percentage import Percentage
from odf.text import P
import openpyxl.styles
from os import remove

try:
    t=gettext.translation('officegenerator',pkg_resources.resource_filename("officegenerator","locale"))
    _=t.gettext
except:
    _=str

def remove_without_errors(filename):
    try:
        remove(filename)
    except OSError as e:
        print(_("Error deleting: {} -> {}".format(filename, e.strerror)))

## If arguments is None, launches with sys.argc parameters. Entry point is toomanyfiles:main
## You can call with main(['--pretend']). It's equivalento to os.system('program --pretend')
## @param arguments is an array with parser arguments. For example: ['--argument','9']. 
def main(arguments=None):
    parser=argparse.ArgumentParser(prog='officegenerator', description=_('Create example files using officegenerator module'), epilog=argparse_epilog(), formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--debug', help="Debug program information", choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")
    group= parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--create', help="Create demo files", action="store_true",default=False)
    group.add_argument('--remove', help="Remove demo files", action="store_true", default=False)
    args=parser.parse_args(arguments)

    addDebugSystem(args.debug)

    if args.remove==True:
        remove_without_errors("officegenerator.ods")
        remove_without_errors("officegenerator.ods.rewritten.ods")
        remove_without_errors("officegenerator.odt")
        remove_without_errors("officegenerator_manual_styles.odt")
        remove_without_errors("officegenerator_search_and_replace.odt")
        remove_without_errors("officegenerator_updated.ods")
        remove_without_errors("officegenerator.xlsx")
        remove_without_errors("officegenerator.xlsx.rewritten.xlsx")
        remove_without_errors("officegenerator.xlsx.data_only.xlsx")
        remove_without_errors("officegenerator_xlsx.data_only.txt")
        remove_without_errors("officegenerator_updated.xlsx")
        remove_without_errors("officegenerator_xlsx_readonly.txt")
        remove_without_errors("officegenerator_ods_readonly.txt")
        remove_without_errors("officegenerator_standard_sheets.ods")
        remove_without_errors("officegenerator_standard_sheets.odt")
        remove_without_errors("officegenerator_standard_sheets.xlsx")
        remove_without_errors("officegenerator_standard_sheets_direct.ods")
        remove_without_errors("officegenerator_standard_sheets_direct.xlsx")

    if args.create==True:
        start=datetime.now()
        futures=[]
        with ProcessPoolExecutor(max_workers=cpu_count()+1) as executor:
            futures.append(executor.submit(demo_ods))
            futures.append(executor.submit(demo_odt_standard))
            futures.append(executor.submit(demo_odt_manual_styles))
            futures.append(executor.submit(demo_xlsx))
            futures.append(executor.submit(demo_standard_sheets))

        for future in as_completed(futures):
            print(future.result())
            
            
        create_rewritten_xlsx("officegenerator.xlsx")
        create_data_only_xlsx("officegenerator.xlsx")
        create_rewritten_ods("officegenerator.ods")
        create_data_only_ods("officegenerator.ods")
        #create_data_only_xlsx("officegenerator.ods", "officegenerator_from_ods.data_only.xlsx")
            
            
        ## This process depend of above files
        futures=[]
        with ProcessPoolExecutor(max_workers=cpu_count()+1) as executor:
            futures.append(executor.submit(demo_ods_updated))
            futures.append(executor.submit(demo_ods_readonly))
            futures.append(executor.submit(demo_odt_search_and_replace))
            futures.append(executor.submit(demo_xlsx_updated))
            futures.append(executor.submit(demo_xlsx_readonly))
            futures.append(executor.submit(demo_xlsx_data_only))

        for future in as_completed(futures):
            print(future.result())
        print("All process took {}".format(datetime.now()-start))

def demo_ods_readonly():
    start=datetime.now()
    doc=ODS_Read("officegenerator.ods")
    output=open("officegenerator_ods_readonly.txt", "w")
    
    output.write("{} {}\n".format("A1",  doc.getCellValue(9, "A1")))        
    output.write("{}\n".format(doc.getColumnValues(1, "J", skip_up=150)))
    output.write("{}\n".format(doc.getRowValues(1, "100", skip_left=3)))
        
    output.write("{}\n".format(doc.values_by_range(9, "A2:K2"))   )
    output.close()
    return "demo_ods_readonly took {}".format(datetime.now()-start)

def demo_ods_updated():
    start=datetime.now()
    doc=ODS_Read("officegenerator.ods")

    #Sustituye celda
    odfcell=doc.getCell(0, "B6")
    odfcell.object=1789.12
    odfcell.setComment(_("This cell has been updated and modified"))
    doc.setCell(0, "B6", odfcell)

    #Added cell
    odfcell=doc.getCell(0, "B10")
    odfcell.object=_("Created cell")
    odfcell.setComment(_("This cell has been updated and modified"))
    doc.setCell(0, "B10", odfcell )

    doc.save("officegenerator_updated.ods")
    return "demo_ods_updated took {}".format(datetime.now()-start)

def demo_ods():
    start=datetime.now()
    doc=ODS_Write("officegenerator.ods")
    doc.setMetadata(_("OfficeGenerator ODS example"),  _("Demo with ODS_Write class"), "Turulomio", _("This file have been generated with OfficeGenerator-{}. You can see OfficeGenerator main page in http://github.com/turulomio/officegenerator").format(__version__), "officegenerator demo files")
    s1=doc.createSheet("Example")
    s1.add("A1", [["Title", "Value"]], "OrangeCenter")
    s1.add("A2", "Percentage", "YellowLeft")
    s1.add("A4",  "Suma", "WhiteCenter")
    s1.add("B2",  Percentage(12, 56), "WhitePercentage")
    s1.add("B3",  Percentage(12, 21), "WhitePercentage")
    s1.add("B4",  "=sum(B2:B3)","WhitePercentage" )
    s1.add("B6",  100.26, "WhiteDecimal6")
    s1.add("B7",  101, "WhiteInteger")
    s1.freezeAndSelect("A2", "A30", "A5")

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


    sf1=doc.createSheet("Freeze A1")
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        for number in range(1, 200):
            sf1.add(letter + str(number), letter+str(number), "YellowLeft")
    sf1.freezeAndSelect("A1", "Z199","U180") 

    sf2=doc.createSheet("Freeze A3")
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        for number in range(1, 200):
            sf2.add(letter + str(number), letter+str(number), "YellowLeft")
    sf2.freezeAndSelect("A3", "Z199","M171") 

    sf3=doc.createSheet("Freeze C1")
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        for number in range(1, 200):
            sf3.add(letter + str(number), letter+str(number), "YellowLeft")
    sf3.freezeAndSelect("C1", "Z199","Q168") 

    sf4=doc.createSheet("Freeze C3")
    for letter in "ABCDEFG":
        for number in range(1, 200):
            sf4.add(letter + str(number), letter+str(number), "YellowLeft")
    sf4.freezeAndSelect("C3", "G199","Q168") 


    sf1=doc.createSheet("Freeze A1 None")
    for letter in "ABCDEFG":
        for number in range(1, 200):
            sf1.add(letter + str(number), letter+str(number), "YellowLeft")
    sf1.freezeAndSelect("A1", "G199") 

    sf2=doc.createSheet("Freeze A3 None")
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        for number in range(1, 15):
            sf2.add(letter + str(number), letter+str(number), "YellowLeft")
    sf2.freezeAndSelect("A3", "Z14") 

    sf3=doc.createSheet("Freeze C1 None")
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        for number in range(1, 15):
            sf3.add(letter + str(number), letter+str(number), "YellowLeft")
    sf3.freezeAndSelect("C1", "Z14") 

    sf4=doc.createSheet("Freeze C3 None")
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        for number in range(1, 200):
            sf4.add(letter + str(number), letter+str(number), "YellowLeft")
    sf4.freezeAndSelect("C3", "Z199") 
    s6=doc.createSheet("Format number")
    s6.setColumnsWidth([ColumnWidthODS.L, ColumnWidthODS.Datetime, ColumnWidthODS.Date, ColumnWidthODS.L, ColumnWidthODS.L, ColumnWidthODS.L, ColumnWidthODS.L, ColumnWidthODS.XL, ColumnWidthODS.L, ColumnWidthODS.L, ColumnWidthODS.L])

    s6.add("A1", _("Style name"), "OrangeCenter")
    s6.add("B1", _("Date and time"), "OrangeCenter")
    s6.add("C1", _("Date"), "OrangeCenter")
    s6.add("D1", _("Integer"), "OrangeCenter")
    s6.add("E1", _("Euros"), "OrangeCenter")
    s6.add("F1", _("Dollars"), "OrangeCenter")
    s6.add("G1", _("Percentage"), "OrangeCenter")
    s6.add("H1", _("Number with 2 decimals"), "OrangeCenter")
    s6.add("I1", _("Number with 6 decimals"), "OrangeCenter")
    s6.add("J1", _("Time"), "OrangeCenter")
    s6.add("K1", _("Boolean"), "OrangeCenter")
    for row, color in enumerate(doc.colors.arr):
        s6.add(Coord("A2").addRow(row), color.name, color.name + "Left")
        s6.add(Coord("B2").addRow(row), datetime.now(),  color.name +"Datetime")
        s6.add(Coord("C2").addRow(row), date.today(), color.name + "Date")
        s6.add(Coord("D2").addRow(row), pow(-1, row)*-10000000, color.name+ "Integer")
        s6.add(Coord("E2").addRow(row), Currency(pow(-1, row)*12.56, "EUR"), color.name + "EUR")
        s6.add(Coord("F2").addRow(row), Currency(pow(-1, row)*12345.56, "USD"), color.name + "USD")
        s6.add(Coord("G2").addRow(row), Percentage(pow(-1, row)*1, 3), color.name+"Percentage")
        s6.add(Coord("H2").addRow(row), pow(-1, row)*123456789.121212, color.name+"Decimal2")
        s6.add(Coord("I2").addRow(row), pow(-1, row)*-12.121212, color.name+"Decimal6")
        s6.add(Coord("J2").addRow(row), (datetime.now()+timedelta(seconds=3600*12*row)).time(), color.name+"Time")
        s6.add(Coord("K2").addRow(row), bool(row%2),  color.name+"Left")

    s6.setComment("B2", _("This is a comment"))
    
    #Merge cells
    s6.addMerged("B13:F14", _("This cell is going to be merged with B13 a F14"), "GreenCenter")
    s6.addMerged("B18:G18", _("This cell is going to be merged and aligned desde B18 a G18"), "YellowRight")
    s6.freezeAndSelect("A11", "B11", "A11")#Default values
    
    #Cells with formula with diferent styles and number formats
    s6.addMerged("A20:D20",  _("These cells show formulas with diferent styles and number formats"), "Green")
    s6.add("E20",  "=2+3.3", "NormalDecimal6")
    s6.add("F20",  "=2+3.3", "YellowEUR")
    s6.add("G20",  "=2+3.3", "GreenDatetime")
    s6.add("H20",  "=0.9/23", "WhitePercentage")


    s7=doc.createSheet("Add totals example")
    s7.add("A1", _("Style name"), "OrangeCenter")
    s7.add("B1", _("Date and time"), "OrangeCenter")
    s7.add("C1", _("Date"), "OrangeCenter")
    s7.add("D1", _("Integer"), "OrangeCenter")
    s7.add("E1", _("Euros"), "OrangeCenter")
    s7.add("F1", _("Dollars"), "OrangeCenter")
    s7.add("G1", _("Percentage"), "OrangeCenter")
    s7.add("H1", _("Number with 2 decimals"), "OrangeCenter")
    s7.add("I1", _("Number with 6 decimals"), "OrangeCenter")
    s7.add("J1", _("Time"), "OrangeCenter")
    s7.add("K1", _("Boolean"), "OrangeCenter")
    for row, color in enumerate(doc.colors.arr):
        s7.add(Coord("A2").addRow(row), color.name, color.name + "Left")
        s7.add(Coord("B2").addRow(row), 1,  color.name +"Integer")
        s7.add(Coord("C2").addRow(row), 1, color.name + "Integer")
        s7.add(Coord("D2").addRow(row), 1, color.name+ "Integer")
        s7.add(Coord("E2").addRow(row), 1, color.name+ "Integer")
        s7.add(Coord("F2").addRow(row), 1, color.name+ "Integer")
        s7.add(Coord("G2").addRow(row), 1, color.name+ "Integer")
        s7.add(Coord("H2").addRow(row), 1, color.name+ "Integer")
        s7.add(Coord("I2").addRow(row), 1, color.name+ "Integer")
        s7.add(Coord("J2").addRow(row), 1, color.name+ "Integer")
        s7.add(Coord("K2").addRow(row), 1, color.name+ "Integer")
    s7.addTotalsHorizontal("A11", ["Total", "#SUM", "#SUM", "#SUM", "#AVG","#MEDIAN","#SUM", "#AVG","#MEDIAN", "#SUM", "#SUM"])
    s7.addTotalsVertical("L1", ["Total", "#SUM", "#AVG","#MEDIAN","#SUM", "#AVG","#MEDIAN", "#SUM", "#SUM", "#SUM", "#SUM"])
    s7.freezeAndSelect("B2") 
    

    doc.setActiveSheet(s6)
    doc.save()

    return "demo_ods took {}".format(datetime.now()-start)
   
def demo_odt_commands(doc):
    doc.setMetadata("OfficeGenerator title",  "OfficeGenerator subject", "Turulomio")
    doc.setMetadata(_("OfficeGenerator ODT example"),  _("Demo with ODT documents"), "Turulomio", _("This file have been generated with OfficeGenerator-{}. You can see OfficeGenerator main page in http://github.com/turulomio/officegenerator").format(__version__), "officegenerator demo files")

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
                            [_("Datetime"), datetime.now()], 
                            [_("Date"), date.today()], 
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
    
    doc.header("Search and replace", 2)
    doc.simpleParagraph("You can search strings in a document and replace them programatically.")
    doc.simpleParagraph("__REPLACEME__")
    doc.simpleParagraph("You can delete this strings if you aren't going to use it")
    doc.simpleParagraph("__DELETEME__")
    
    doc.header("ODS Reading", 1)
    doc.pageBreak()

    doc.header("XLSX", 1)
    
def demo_odt_search_and_replace():
    start=datetime.now()
    doc=ODT("officegenerator_search_and_replace.odt", template="officegenerator.odt")
    doc.search_and_replace("_DELETEME__", "")
    doc.search_and_replace("__REPLACEME__",  "This text has been replaced programatically")
    doc.save()
    return "demo_odt_search_and_replace took {}".format(datetime.now()-start)

    
def demo_odt_standard():
    start=datetime.now()
    doc=ODT_Standard("officegenerator.odt")
    demo_odt_commands(doc)
    doc.save()
    return "demo_odt_standard took {}".format(datetime.now()-start)
    
def demo_odt_manual_styles():
    start=datetime.now()
    doc=ODT_Manual_Styles("officegenerator_manual_styles.odt")
    demo_odt_commands(doc)
    doc.save()
    return "demo_odt_manual_styles took {}".format(datetime.now()-start)

def demo_xlsx():
    start=datetime.now()
    xlsx=XLSX_Write("officegenerator.xlsx")
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
    xlsx.overwrite("I1", _("Time"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("J1", _("Boolean"), style=xlsx.stOrange,  alignment="center")
    for row, style in enumerate([xlsx.stOrange, xlsx.stGreen, xlsx.stGrayLight, xlsx.stYellow, xlsx.stGrayDark, xlsx.stWhite, None]):
        xlsx.overwrite(Coord("A2").addRow(row), xlsx.styleName(style), style=style)
        xlsx.overwrite(Coord("B2").addRow(row), datetime.now(), style=style)
        xlsx.overwrite(Coord("C2").addRow(row), date.today(), style=style)
        xlsx.overwrite(Coord("D2").addRow(row), pow(-1, row)*-10000000, style=style)
        xlsx.overwrite(Coord("E2").addRow(row), Currency(pow(-1, row)*12.56, "EUR"), style=style)
        xlsx.overwrite(Coord("F2").addRow(row), Percentage(1, 3), style=style,  decimals=row+1)
        xlsx.overwrite(Coord("G2").addRow(row), pow(-1, row)*12.121212, style=style, decimals=2)
        xlsx.overwrite(Coord("H2").addRow(row), pow(-1, row)*-12.121212, style=style, decimals=6)
        xlsx.overwrite(Coord("I2").addRow(row), datetime.now().time(), style=style, decimals=6)
        xlsx.overwrite(Coord("J2").addRow(row), True, style=style, decimals=6)
    xlsx.setComment("B2", _("This is a comment"))
    
    ##To write a custom cell
    cell=xlsx.wb.active['B12']
    cell.font=openpyxl.styles.Font(name='Arial', size=16, bold=True, color=openpyxl.styles.colors.RED)
    cell.value=_("This is a custom cell")
    #Merge cells
    xlsx.overwrite_and_merge("A13:C14", _("This cell is going to be merged with B13 and C13"),style=xlsx.stOrange)
    xlsx.overwrite_and_merge("A18:G18", _("This cell is going to be merged and aligned"),style=xlsx.stGrayDark, alignment="right")

    xlsx.overwrite("A16", [[_("This are booleans"), False, True]], style=xlsx.stWhite, alignment='right')
    xlsx.overwrite("D16", _("These are formulas returning booleans"), style=xlsx.stGreen)

    xlsx.overwrite("A20",  [["Una fila"]*3], style=xlsx.stGrayDark)
    xlsx.overwrite_and_merge("E13:G13", _("This sheet max rows are {} and max columns {}").format(xlsx.rowNumber(xlsx.ws_current_id), xlsx.columnNumber(xlsx.ws_current_id)), style=xlsx.stYellow,  alignment="center")

    #Named cells
    xlsx.overwrite_and_merge("A23:B23", _("Cell B23 has a name 'Amount"), style=xlsx.stWhite)
    xlsx.overwrite("C23", 5, style=xlsx.stWhite)
    xlsx.setCellName("$C$23", "Amount")

    xlsx.overwrite_and_merge("A24:B24", _("Cell B24 has a name 'Price"), style=xlsx.stWhite)
    xlsx.overwrite("C24", Currency(10,'EUR'), style=xlsx.stWhite)
    xlsx.setCellName("$C$24", "Price")

    xlsx.overwrite_and_merge("A25:B25", _("Cell B25 has a product with names"), style=xlsx.stWhite)
    xlsx.overwrite_formula("C25", "=Amount*Price", "€", style=xlsx.stWhite, alignment='right')

    xlsx.freezeAndSelect("A9","B11", "A9")
   
    
    #To text split and cur position
    xlsx.createSheet("Freeze A1 None")
    for letter in "ABCDEFGHIJKLMNOPQ":
        for number in range(1, 100):
            xlsx.overwrite(letter + str(number), letter+str(number), style=xlsx.stYellow)
    xlsx.freezeAndSelect("A1","Z199")
    #To text split and cur position
    xlsx.createSheet("Freeze A3 None")
    for letter in "ABCDEFGHIJKLMNOPQ":
        for number in range(1, 100):
            xlsx.overwrite(letter + str(number), letter+str(number), style=xlsx.stYellow)
    xlsx.freezeAndSelect("A3","Z199")
    #To text split and cur position
    xlsx.createSheet("Freeze C1 None")
    for letter in "ABCDEFGHIJKLMNOPQ":
        for number in range(1, 100):
            xlsx.overwrite(letter + str(number), letter+str(number), style=xlsx.stYellow)
    xlsx.freezeAndSelect("C1","Z199")
    #To text split and cur position
    xlsx.createSheet("Freeze C3 None")
    for letter in "ABCDEFGHIJKLMNOPQ":
        for number in range(1, 100):
            xlsx.overwrite(letter + str(number), letter+str(number), style=xlsx.stYellow)
    xlsx.freezeAndSelect("C3","Z199")
    
    xlsx.createSheet("Add totals example")
    xlsx.overwrite("A1", _("Style name"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("B1", _("Date and time"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("C1", _("Date"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("D1", _("Integer"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("E1", _("Euros"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("F1", _("Percentage"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("G1", _("Number with 2 decimals"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("H1", _("Number with 6 decimals"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("I1", _("Time"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("J1", _("Boolean"), style=xlsx.stOrange,  alignment="center")
    xlsx.overwrite("K1", _("Boolean"), style=xlsx.stOrange,  alignment="center")
    for row, style in enumerate([xlsx.stOrange, xlsx.stGreen, xlsx.stGrayLight, xlsx.stYellow, xlsx.stGrayDark, xlsx.stWhite, None]):
        xlsx.overwrite(Coord("A2").addRow(row), xlsx.styleName(style), style=style)
        xlsx.overwrite(Coord("B2").addRow(row), 1, style=style)
        xlsx.overwrite(Coord("C2").addRow(row), 1, style=style)
        xlsx.overwrite(Coord("D2").addRow(row), 1, style=style)
        xlsx.overwrite(Coord("E2").addRow(row), 1, style=style)
        xlsx.overwrite(Coord("F2").addRow(row), 1, style=style,  decimals=row+1)
        xlsx.overwrite(Coord("G2").addRow(row), 1, style=style, decimals=2)
        xlsx.overwrite(Coord("H2").addRow(row), 1, style=style, decimals=6)
        xlsx.overwrite(Coord("I2").addRow(row), 1, style=style, decimals=6)
        xlsx.overwrite(Coord("J2").addRow(row), 1, style=style, decimals=6)
        xlsx.overwrite(Coord("K2").addRow(row), 1, style=style, decimals=6)
    xlsx.overwriteTotalsHorizontal("A9", ["Total", "#SUM", "#SUM", "#SUM", "#AVG","#MEDIAN","#SUM", "#AVG","#MEDIAN", "#SUM", "#SUM"])
    xlsx.overwriteTotalsVertical("L1", ["Total", "#SUM", "#AVG","#MEDIAN","#SUM", "#AVG","#MEDIAN", "#SUM","#SUM"])
    xlsx.freezeAndSelect("B2") 
    
    xlsx.save()
    return "demo_xlsx took {}".format(datetime.now()-start)

def demo_xlsx_updated():
    start=datetime.now()
    xlsx=XLSX_Write("officegenerator_updated.xlsx", "officegenerator.xlsx")
    xlsx.setCurrentSheet(0)
    
    xlsx.overwrite("A2", _("Orange"))
    xlsx.overwrite("A5", _("Yellow"))
    xlsx.overwrite("A7", _("White"), style=xlsx.stWhite, alignment="center")
    xlsx.overwrite("D4", 1223)
    #Merge cells
    xlsx.overwrite_and_merge("A15:C16", _("This cell is going to be merged with B13 and C13"),style=xlsx.stOrange)
    xlsx.overwrite_and_merge("A17:G17", _("This cell is going to be merged and aligned"),style=xlsx.stGrayDark, alignment="right")

    xlsx.save()
    return "demo_xlsx_updated took {}".format(datetime.now()-start)
    
def demo_xlsx_readonly(): 
    start=datetime.now()
    doc=XLSX_Read("officegenerator.xlsx")
    output=open("officegenerator_xlsx_readonly.txt", "w")
        
    range_=Range("A2:J2")
    for coord in range_.coords()[0]:
        output.write("{} {}\n".format(coord,  doc.getCellValue(0, coord)))
        
    output.write("{}\n".format(doc.getColumnValues(0, "J", skip_up=0)))
    output.write("{}\n".format(doc.getRowValues(1, "100", skip_left=3)))
        
    output.write("{}\n".format(doc.values_by_range(0, range_) ))
    output.close()
    return "demo_xlsx_readonly took {}".format(datetime.now()-start)
        
    
def demo_xlsx_data_only():
    start=datetime.now()   
    output=open("officegenerator_xlsx.data_only.txt", "w")

    # To read a XLSX_Write created file with data_only fails Values are not generated by libreoffice, just oopenpyxl
    doc=XLSX_Read("officegenerator.xlsx", data_only=True)
    output.write(str(doc.getCellValue(0, "C25"))+"\n")
    #File with formulas. Works. Openpyxl generated formulas
    doc=XLSX_Read("officegenerator.xlsx")
    output.write(str(doc.getCellValue(0, "C25"))+"\n")
    #File converted to data_only. Works to read value due libreoffice updated values in create_data_only_filename function
    doc=XLSX_Read("officegenerator.xlsx.rewritten.xlsx")
    output.write(str(doc.getCellValue(0, "C25"))+"\n")
    
    output.close()
    return "demo_xlsx_data_only took {}".format(datetime.now()-start)
    
def demo_standard_sheets():
    start=datetime.now()
    ods=ODS_Write("officegenerator_standard_sheets.ods")
    odt=ODT_Standard("officegenerator_standard_sheets.odt")
    xlsx=XLSX_Write("officegenerator_standard_sheets.xlsx")
    
    m=Model()
    m.setTitle("HV")
    m.setHorizontalHeaders(["Number", "Data", "More data", "Dt"], [1, 2, 3, 4])
    m.setVerticalHeaders(["V1", "V2", "V3"]*10, 4)
    data=[]        
    for row in range(30):
        data.append([row, "Data", "Data++", datetime.now()+timedelta(days=row)])
    m.setData(data)
    m.ods_sheet(ods)
    m.xlsx_sheet(xlsx)
    m.odt_table(odt, 15, 8)
    
    m2=Model()
    m2.setTitle("V")
    m2.setHorizontalHeaders(None, [1, 2, 3])
    m2.setVerticalHeaders(["Number", "Data", "More data"]*10)
    m2.setData(data)
    m2.removeColumns([1, 2, ])
    m2.removeRows([1, 2, ])
    m2.ods_sheet(ods)
    m2.xlsx_sheet(xlsx)
    m2.odt_table(odt, 15, 10)
    
    m=Model()
    m.setTitle("H totals")
    m.setHorizontalHeaders(["Concept", "Decimal", "Currency", "Percentage"], [5, 3, 3, 3])
    data=[]        
    for row in range(30):
        data.append([f"Concept {row}", row*10, Currency(row*10/7, "EUR"), Percentage(row, 12) ])
    m.setData(data)
    m.setHorizontalTotalDefinition(["Total", "#SUM","#AVG","#MEDIAN" ])
    m.ods_sheet(ods)
    m.xlsx_sheet(xlsx)
    m.odt_table(odt, 15, 8)    
    m=Model()

    m.setTitle("H totals skip 2")
    m.setHorizontalHeaders(["Concept", "Decimal", "Currency", "Percentage"], [5, 3, 3, 3])
    data=[]        
    for row in range(30):
        data.append([f"Concept {row}", row*10, Currency(row*10/7, "EUR"), Percentage(row, 12) ])
    m.setData(data)
    m.setHorizontalTotalDefinition(["Total", "#SUM","#AVG","#MEDIAN" ],totals_index_from=2)
    m.ods_sheet(ods)
    m.xlsx_sheet(xlsx)
    m.odt_table(odt, 15, 8)
    
    m=Model()
    m.setTitle("V totals")
    m.setHorizontalHeaders(["Concept", "Decimal", "Decimal2", "Decimal3"], [5, 3, 3, 3])
    data=[]        
    for row in range(30):
        data.append([f"Concept {row}", row*10, row*10, row*10 ])
    m.setData(data)
    m.setVerticalTotalDefinition(["Total"]+["#SUM"]*m.numDataRows() )
    m.ods_sheet(ods)
    m.xlsx_sheet(xlsx)
    m.odt_table(odt, 15, 8)    
    m=Model()

    m.setTitle("HV totals")
    hh=["Concept", "Decimal", "Decimal2", "Decimal3"]
    m.setHorizontalHeaders(hh, [5, 3, 3, 3])
    data=[]        
    for row in range(30):
        data.append([f"Concept {row}", row*10, row*10, row*10 ])
    m.setData(data)
    m.setHorizontalTotalDefinition(["Total", "#SUM","#AVG","#MEDIAN" ])
    m.setVerticalTotalDefinition(["Total"]+["#SUM"]*(m.numDataRows() +1))
    m.ods_sheet(ods)
    m.xlsx_sheet(xlsx)
    m.odt_table(odt, 15, 8)
    
    #Model auto
    hh=["Concept", "Decimal", "Decimal2", "Decimal3"]
    data=[]        
    for row in range(30):
        data.append([f"Concept {row}", row*10, -row*10, row*10 ])
    m=Model_Auto("Model_Auto", hh, data)
    m.order_with_none(1, reverse=True)
    m.ods_sheet(ods)
    m.xlsx_sheet(xlsx)
    m.odt_table(odt, 15, 8)
    
    
    #Model auto
    hh=["Concept", "Concept2", "Decimal", "Decimal2", "Decimal3"]
    data=[]        
    for row in range(30):
        data.append([f"Concept {row}", datetime.now(),  Currency(row*10, 'EUR'), Currency(-row*10, 'EUR'), Currency(row*10, 'EUR') ])
    m=Model_Auto("Model_Auto skip 2", hh, data, column_index_from=2)
    m.ods_sheet(ods)
    m.xlsx_sheet(xlsx)
    m.odt_table(odt, 15, 8)
    
    ods.save()
    xlsx.save()
    odt.save()
    
    m.ods_file("officegenerator_standard_sheets_direct.ods")
    m.xlsx_file("officegenerator_standard_sheets_direct.xlsx")
    return "demo_standard_sheets took {}".format(datetime.now()-start)


if __name__ == "__main__":
    main()
