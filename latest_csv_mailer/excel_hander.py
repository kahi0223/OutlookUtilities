from config import ExcelConfig
from win32com.client import DispatchEx
import traceback

class ExcelHander:
    REFRESH_MACRO = 'MacroName'
    def refresh_excel(self, excel_path:str) -> bool:
        print("----Refreshing Excel")
        try:
            excel = DispatchEx("Excel.Application")
            excel.Visible = 1
            excel.Workbooks.Open(Filename = excel_path , ReadOnly = 0)
            excel.Application.Run("'" + excel_path + "'" + REFRESH_MACRO)
            print("> Refreshed")
            excel.Workbooks(1).Close(SaveChanges = 1)
            excel.Application.Quit()
            return True
        except:
            print("?? Refresh 2 Error")
            print(traceback.format_exc())  
            return False