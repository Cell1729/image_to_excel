import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime
import os


class ImportExcel():
    """
    prepare the excel file for further processing
    """
    def __init__(self):
        self.Settingtemplates()
        self.CopyFromtemplates()

    def Settingtemplates(self):
        templates_path = os.path.join(os.path.dirname(__file__), '..', 'templates', 'templates.xlsx')
        self.templates_excel = openpyxl.load_workbook(templates_path)
        self.templates_sheet = self.templates_excel['templates']
    
    def CopyFromtemplates(self):
        now = datetime.now()
        self.filename = f'{now:%Y-%m-%d-%H-%M}'
        
        output_dir = os.path.join(os.path.dirname(__file__), '..', 'output')
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        self.output_path = os.path.join(output_dir, f'{self.filename}.xlsx')
        
        self.templates_excel.save(self.output_path)
    
    def GetExcel(self):
        return self.filename, self.output_path

class EditExcel():
    def __init__(self):
        importexcel = ImportExcel()
        self.filename, self.Excel_path = importexcel.GetExcel()
        self.OutPutExcel = openpyxl.load_workbook(self.Excel_path)
        self.OutPutSheet = self.OutPutExcel['templates']
    
    def fillColor(self, x, y, color):
        try:
            self.OutPutSheet.cell(row=x+1, column=y+1).fill = PatternFill(patternType='solid', fgColor=color, bgColor=color)
        except Exception as e:
            print(f"An error occurred: {e}")

    def saveExcel(self):
        output_path = os.path.join(os.path.dirname(__file__), '..', 'output', f'{self.filename}.xlsx')
        self.OutPutExcel.save(output_path)
        print('success!!')