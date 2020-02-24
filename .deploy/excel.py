import win32com.client as win32 
from pywintypes import com_error 
class Excel:
  def __init__(self):
    pass 
  
  def __enter__(self, visible = True):
    self.app = win32.Dispatch('Excel.Application')
    self.app.Visible = visible 
    return self
  
  def __exit__(self, type, value, trace):
    self.book.Close() if self.book else 0
    self.app.Quit()
  
  def open(self, path:str):
    self.book = self.app.Workbooks.Open(path)
    return self 
  
  def to_pdf(self, path:str):
    try:
      self.book.Worksheets.Select()
      self.book.ActiveSheet.ExportAsFixedFormat(0, path)
    except com_error as e:
      print(e)
    return self
  