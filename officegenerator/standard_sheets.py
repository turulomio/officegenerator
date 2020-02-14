from officegenerator.commons import Coord

class Model:
    def __init__(self):
        self.hh_ods_style="OrangeCenter"
        self.vh_ods_style="GreenLeft"
        self.hh=None
        self.vh=None

    def setTitle(self, title):
        self.title=title
        
    ## If you want only vertical headers you must add self.setHorizontalHeaders(None, sizes=10)
    ## @param hh List or NOne
    ## @param sizes int or List in cm, to generate document use, ods_columnSizes
    def setHorizontalHeaders(self, hh, sizes=10):
        self.hh=hh
        if hh==None:#Only vertical headers
            columns=100
        else:
            columns=len(self.hh)
        if sizes.__class__.__name__=="int":
            self.hh_sizes=[sizes]*columns
        else:
            self.hh_sizes=sizes
    
    def setVerticalHeaders(self, vh):
        self.vh=vh
        
    def ods_columnSizes(self):
        r=[]
        for arg in self.hh_sizes:
                r.append(arg*30)
        return r
        
    def xlsx_columnSizes(self):
        r=[]
        for arg in self.hh_sizes:
                r.append(arg*6)
        return r

    def setData(self, data):
        self.data=data
           
            
    ## @param title String with the title of the sheet
    ## @param columns_title List of Strings
    ## @param data list of list
    ## @param sizes List of integers
    def xlsx_sheet(self, doc):
        doc.createSheet(self.title)
        doc.setColumnsWidth(self.xlsx_columnSizes())
        if self.hh is not None:
            doc.overwrite(self.__getFirstContentCoord().addRow(-1), [self.hh], doc.stOrange)
        if self.vh is not None:
            for i, header in enumerate(self.vh):
                doc.overwrite(self.__getFirstContentCoord().addColumn(-1).addRow(i), header, doc.stGreen, alignment="left")
        for number, row in enumerate(self.data):
            for letter,  field in enumerate(row):
                doc.overwrite(self.__getFirstContentCoord().addRow(number).addColumn(letter), field)
        doc.freezeAndSelect(self.__getFirstContentCoord(), self.__getFirstContentCoord().addRow(number).addColumn(letter))

    ## @param title String with the title of the sheet
    ## @param columns_title List of Strings
    ## @param data list of list
    ## @param sizes List of integers
    def ods_sheet(self, doc):
        s=doc.createSheet(self.title)
        s.setColumnsWidth(self.ods_columnSizes())
        if self.hh is not None:
            s.add(self.__getFirstContentCoord().addRow(-1), [self.hh], self.hh_ods_style)
        if self.vh is not None:
            for i, header in enumerate(self.vh):
                s.add(self.__getFirstContentCoord().addColumn(-1).addRow(i), header, self.vh_ods_style)
                
        for number, row in enumerate(self.data):
            for letter,  field in enumerate(row):
                s.add(self.__getFirstContentCoord().addRow(number).addColumn(letter), field)
        s.freezeAndSelect(self.__getFirstContentCoord(),self.__getFirstContentCoord().addRow(number).addColumn(letter))

    def odt_table(self, doc, sizes, fontsize, after=True):
        return doc.table(self.hh, self.data, sizes, fontsize, self.title, after)        
        
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

if __name__ == "__main__":
    from officegenerator.libodfgenerator import  ODS_Write
    from officegenerator.libodfgenerator import ODT_Standard
    from officegenerator.libxlsxgenerator import OpenPyXL
    filename="standard_sheets.ods"
    ods=ODS_Write("standard_sheets.ods")
    odt=ODT_Standard("standard_sheets.odt")
    xlsx=OpenPyXL("standard_sheets.xlsx")
    
    m=Model()
    m.setTitle("Probe")
    m.setHorizontalHeaders(["Number", "Data", "More data"], 7)
    m.setVerticalHeaders(["Number", "Data", "More data"]*10)
    data=[]        
    for row in range(30):
        data.append([row, "Data", "Data++"])
    m.setData(data)
    m.ods_sheet(ods)
    m.xlsx_sheet(xlsx)
    
    m2=Model()
    m2.setTitle("Probe 2")
    m2.setHorizontalHeaders(None, 7)
    m2.setVerticalHeaders(["Number", "Data", "More data"]*10)
    m2.setData(data)
    m2.ods_sheet(ods)
    m2.xlsx_sheet(xlsx)
    
    m.odt_table(odt, [3]*3, 8)
    
    xlsx.remove_sheet_by_id(0)
    ods.save()
    xlsx.save()
    odt.save()
