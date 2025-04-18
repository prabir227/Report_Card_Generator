from docx.shared import Pt
from docx import Document
from openpyxl import load_workbook
from docx.enum.text import WD_ALIGN_PARAGRAPH
class ExcelOperations:
    def __init__(self,lvl0Start, lvl0End, lvl1Start, lvl1End, lvl2Start, lvl2End, lvl3Start, lvl3End, lvl4Start, lvl4End, lvl5Start, lvl5End, eng1="C", eng2="K", eng3="S", hin1="D", hin2="L", hin3="T", math1="E", math2="M", math3="U", sst1="G", sst2="O", sst3="W", evs1="F", evs2="N", evs3="V", totalLvl4=150, totalLvl5=200 , centre="Boring Road"):
        self.__eng1 = eng1
        self.__eng2 = eng2
        self.__eng3 = eng3
        self.__hin1 = hin1
        self.__hin2 = hin2
        self.__hin3 = hin3
        self.__math1 = math1
        self.__math2 = math2
        self.__math3 = math3
        self.__sst1 = sst1
        self.__sst2 = sst2
        self.__sst3 = sst3
        self.__evs1 = evs1
        self.__evs2 = evs2
        self.__evs3 = evs3
        self.__totalLvl4 = totalLvl4
        self.__totalLvl5 = totalLvl5
        self.__centre = centre
        self.__lvl0Start = lvl0Start
        self.__lvl0End = lvl0End
        self.__lvl1Start = lvl1Start
        self.__lvl1End = lvl1End
        self.__lvl2Start = lvl2Start
        self.__lvl2End = lvl2End
        self.__lvl3Start = lvl3Start
        self.__lvl3End = lvl3End
        self.__lvl4Start = lvl4Start
        self.__lvl4End = lvl4End
        self.__lvl5Start = lvl5Start
        self.__lvl5End = lvl5End
        self.__doc = Document("resultFormat.docx")
        self.__result = load_workbook("uploads/uploaded.xlsx")
        self.__marksTable = self.__doc.tables[1]
        self.__nameTable = self.__doc.tables[0]
    def docxWriter(self):

        

