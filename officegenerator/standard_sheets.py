## DEBE INCLUIRSE IN OFFICEGENERATOR
from officegenerator.commons import Coord, rowAdd
from officegenerator.libxlsxgenerator import OpenPyXL

class Model:
    def __init__(self):
        self.columns_sizes=None
        self.title_style=None

    def setTitle(self, title):
        self.title=title
        
    def setColumns(self, columns_title):
        self.columns=columns_title
        
    ## Number as cm
    ## 
    def setColumnSizes(self, *args):
        self.columns_sizes=args
        
    def ods_columnSizes(self):
        r=[]
        for arg in self.columns_sizes:
                r.append(arg*30)
        return r
        
    def xlsx_columnSizes(self):
        r=[]
        for arg in self.columns_sizes:
                r.append(arg*6)
        return r

    def setData(self, data):
        self.data=data
        
    def setTitleStyle(self, s):
        self.title_style=s

    def odsTitleStyle(self):
            if self.title_style==None:
                return "OrangeCenter"

    def xlsxTitleStyle(self, doc):
            if self.title_style==None:
                return doc.stOrange
            
            
    ## @param title String with the title of the sheet
    ## @param columns_title List of Strings
    ## @param data list of list
    ## @param sizes List of integers
    def xlsx_sheet(self, doc):
        doc.createSheet(self.title)
        doc.setColumnsWidth(self.xlsx_columnSizes())
        doc.overwrite("A1", [self.columns], self.xlsxTitleStyle(doc))
        for number, row in enumerate(self.data):
            for letter,  field in enumerate(row):
                doc.overwrite(Coord("A2").addRow(number).addColumn(letter), field)
        doc.freezeAndSelect("A2", "A{}".format(rowAdd("2", number)), "A{}".format(rowAdd("2", number-20)))

    ## @param title String with the title of the sheet
    ## @param columns_title List of Strings
    ## @param data list of list
    ## @param sizes List of integers
    def ods_sheet(self, doc):
        s=doc.createSheet(self.title)
        s.setColumnsWidth(self.ods_columnSizes())
        s.add("A1", [self.columns], self.odsTitleStyle())
        for number, row in enumerate(self.data):
            for letter,  field in enumerate(row):
                s.add(Coord("A2").addRow(number).addColumn(letter), field)
        s.setSplitPosition("A1")
        s.setCursorPosition(Coord("B2").addRow(len(self.data)))

if __name__ == "__main__":
    from officegenerator.libodfgenerator import  ODS_Write
    filename="standard_sheets.ods"
    ods=ODS_Write("standard_sheets.ods")
    xlsx=OpenPyXL("standard_sheets.xlsx")
    m=Model()
    m.setTitle("Probe")
    m.setColumnSizes(1, 2, 3)
    m.setColumns(["Number", "Data", "More data"])
    data=[]        
    for row in range(30):
        data.append([row, "Data", "Data++"])
    m.setData(data)
    m.ods_sheet(ods)
    m.xlsx_sheet(xlsx)
    ods.save()
    xlsx.save()
