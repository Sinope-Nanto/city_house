import openpyxl
from openpyxl.styles import Side, Border,PatternFill, colors, Font, Alignment
from openpyxl.drawing.image import Image
import time

from docx.oxml.ns import qn
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt,Cm


class GenExcelReport90:
    def __init__(self,url):
        self.edge = Side(style='thin', color='000000')
        self.border = Border(top=self.edge, bottom=self.edge, left=self.edge, right=self.edge)
        self.url = url
        self.report = openpyxl.Workbook()
    def create_report(self):
        indexsummary = self.report.active
        indexsummary.title = '指数汇总'
        totalradioplot = self.report.create_sheet()
        totalradioplot.title = '全国指数同比环比图'
        plotsimple = self.report.create_sheet()
        plotsimple.title = '作图（简版）'
        plotcomplex = self.report.create_sheet()
        plotcomplex.title = '作图（详细）'
        citysummary = self.report.create_sheet()
        citysummary.title = '城市汇总'
        cityline = self.report.create_sheet()
        cityline.title = '各线城市汇总'
        citysort = self.report.create_sheet()
        citysort.title = '城市排序'
        cityindex = self.report.create_sheet()
        cityindex.title = '各城市指数'
        cityvolumn = self.report.create_sheet()
        cityvolumn.title = '各城市样本量'
        cityplot = self.report.create_sheet()
        cityplot.title = '各城市作图'
        line = self.report.create_sheet()
        line.title = '各线城市子市场'

        return 
    def EndReport(self):
        self.report.save(self.url)

    def to_date(self, temp):
        year = int(temp)//12 + 6
        month = int(temp)%12 + 1
        return str(year).zfill(2) + '-' + str(month).zfill(2)

class GenWordReport90:
    def __init__(self,url):
        self.url = url
        self.doc = Document()
    
    def EndReport(self):
        self.doc.save(self.url)

class GenWordPicture90:

    def __init__(self,url):
        self.url = url
        self.doc = Document()    

    def EndReport(self):
        self.doc.save(self.url)