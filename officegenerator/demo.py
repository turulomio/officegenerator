## @namespace officegenerator.demo
## @brief Generate ODF example files
##
## Prueba

import argparse
import datetime
import os

from .__init__ import __version__, __versiondate__
from .libodfgenerator import ODS_Read, ODS_Write, ODT, OdfCell, OdfPercentage, OdfMoney, rowAdd, columnAdd
from odf.text import P

def main():
    parser=argparse.ArgumentParser(prog='officegenerator', description='Create example files using officegenerator module', epilog="Developed by Mariano Muñoz 2018-{}".format(__versiondate__.year), formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version=__version__)
    group= parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--create', help="Create demo files", action="store_true",default=False)
    group.add_argument('--remove', help="Remove demo files", action="store_true", default=False)
    args=parser.parse_args()

    if args.remove==True:
        os.system("rm libodfgenerator.ods")
        os.system("rm libodfgenerator.odt")
        os.system("rm libodfgenerator_readed.ods")

    if args.create==True:
        #ODS
        doc=ODS_Write("libodfgenerator.ods")
        doc.setMetadata("LibODFGenerator example",  "This class documentation", "Mariano Muñoz")
        s1=doc.createSheet("Example")
        s1.add("A", "1", [["Title", "Value"]], "HeaderOrange")
        s1.add("A", "2", "Percentage", "TextLeft")
        s1.add("A", "4",  "Suma", "TextRight")
        s1.add("B", "2",  OdfPercentage(12, 56))
        s1.add("B", "3",  OdfPercentage(12, 56))
        s1.add("B", "4",  "=sum(B2:B3)","Percentage" )
        s1.add("B", "6",  100.26)
        s1.add("B", "7",  101)
        s1.setCursorPosition("A", "3")
        s1.setSplitPosition("A", "2")

        s2=doc.createSheet("Example 2")
        s2.add("A", "1", [["Title", "Value"]], "HeaderOrange")
        s2.add("A", "2", "Currency", "TextLeft")
        s2.add("B", "2",  OdfMoney(12, "EUR"))
        s2.add("A", "3", "Datetime", "TextLeft")
        s2.add("B", "3",  datetime.datetime.now())
        s2.add("A", "4", "Date", "TextLeft")
        s2.add("B", "4",  datetime.date.today())
        s2.setColumnsWidth([330, 150])
        s2.setCursorPosition("D", "6")
        s2.setSplitPosition("B", "2")
        
        #Adding a cell to s1 after s2 definition
        cell=OdfCell("B", "10", "Celda con OdfCell", "HeaderYellow")
        cell.setComment("Comentario")
        cell.setSpanning(2, 2)
        s1.addCell(cell)
        
        s3=doc.createSheet("Styles")
        s3.setColumnsWidth([400, 150, 150])
        s3.add("A","1","LibODFGenerator has the folowing default Styles:")
        for number,  style in enumerate(["HeaderOrange", "HeaderYellow", "HeaderGreen", "HeaderRed", "HeaderGray", "HeaderOrangeLeft", "HeaderYellowLeft","HeaderGreenLeft",  "HeaderGrayLeft", "TextLeft", "TextRight", "TextCenter"]):
            s3.add("B", rowAdd("1", number) , style, style=style)
        s3.add("A",rowAdd("2", number+1) ,"LibODFGenerator has the folowing default cell classes:")
        s3.add("B",rowAdd("2", number+1) ,OdfMoney(1234.23, "EUR"))
        s3.add("C",rowAdd("2", number+1) ,OdfMoney(-1234.23, "EUR"))
        s3.add("B",rowAdd("2", number+2) ,OdfPercentage(1234.23, 10000))
        s3.add("C",rowAdd("2", number+2) ,OdfPercentage(-1234.23, 25000))
    
    
        s4=doc.createSheet("Splitting")
        for letter in "ABCDEFGHIJ":
            for number in range(1, 11):
                s4.add(letter, str(number), letter+str(number), "HeaderYellowLeft")
        s4.setCursorPosition("C", "3")
        s4.setSplitPosition("C", "3")
        
        doc.setActiveSheet(s3)
        doc.save()
        print("ODS Generated")
    
        doc=ODS_Read("libodfgenerator.ods")
        s1=doc.getSheetElementByIndex(0)
        print("Getting values from ODS:")
        print("  + String", doc.getCellValue(s1, "A", "1"))
        print("  + Percentage", doc.getCellValue(s1, "B", "2"))
        print("  + Formula", doc.getCellValue(s1, "B", "4"))
        print("  + Decimal", doc.getCellValue(s1, "B", "6"))
        print("  + Decimal", doc.getCellValue(s1, "B", "7"))
        s2=doc.getSheetElementByIndex(1)
        print("  + Currency", doc.getCellValue(s2, "B", "2"))
        print("  + Datetime", doc.getCellValue(s2, "B", "3"))
        print("  + Date", doc.getCellValue(s2, "B", "4"))
        
        ##Sustituye celda
        odfcell=doc.getCell(s1, "B", "6")
        odfcell.object=1789.12
        doc.setCell(s1, "B", "6", odfcell)
        doc.save("libodfgenerator_readed.ods")
    
        odfcell=doc.getCell(s1, "B", "10")
        odfcell.object="TURULETE"
    #    odfcell.setComment("Turulete")
        doc.setCell(s1, "B", "10", odfcell )
    
        #ODT#
        doc=ODT("libodfgenerator.odt", language="fr", country="FR")
        doc.setMetadata("LibODFGenerator manual",  "LibODFGenerator documentation", "Mariano Muñoz")
        doc.title("Manual of LibODFGenerator")
        doc.header("ODT Writing", 1)
        doc.simpleParagraph("Hola a todos")
        doc.list(["Pryueba hola no", "Adios", "Bienvenido"], style="BulletList")
        doc.simpleParagraph("Hola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todosHola a todos")
        doc.numberedList(["Pryueba hola no", "Adios", "Bienvenido"])
        doc.simpleParagraph("Con libodfgenerator podemos")
        doc.simpleParagraph("This library create several default styles for writing ODT files:")
        doc.list(["Title: Generates a title with 18pt and bold font", "Header1: Generates a Level 1 header"], style="BulletList")
        doc.addImage("/usr/share/officegenerator/images/crown.png","images/crown.png")
        p = P(stylename="Standard")
        p.addText("Este es un ejemplo de imagen as char: ")
        p.addElement(doc.image("images/crown.png", "3cm", "3cm"))
        p.addText(". Ahora sigo escribiendo sin problemas.")
        doc.doc.text.addElement(p)
        doc.simpleParagraph("Como ves puedo repetirla mil veces sin que me aumente el tamaño del fichero, porque uso referencias")
        p=P(stylename="Standard")
        for i in range(100):
            p.addElement(doc.image("images/crown.png", "4cm", "4cm"))
        p.addText(". Se acabó.")
        doc.doc.text.addElement(p)
        doc.pageBreak()

        doc.header("ODS Writing", 1)
        doc.simpleParagraph("This library create several default styles for writing ODS files. You can see examples in libodfgenerator.ods.")
        doc.pageBreak(horizontal=True)

        doc.header("ODS Reading", 1)
        doc.save()
        print("ODT Generated")

if __name__ == "__main__":
    main()