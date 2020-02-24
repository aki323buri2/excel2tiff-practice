import sys 
from pathlib import Path 
# args
xlsx_path = '011_納品書_タテ型.xlsx'
xlsx_path = r'D:\data\dev\drill4\app\backend\py\011_納品書_タテ型.xlsx'
# args

def excel2pdf(xlsx_path):
  xlsx = Path(xlsx_path)
  root = Path(__file__).parent 
  xlsx = xlsx if xlsx.is_absolute() else root / xlsx

  pdf_dir = root / 'pdf'
  pdf_dir.mkdir(parents=True, exist_ok=True)
  pdf = (pdf_dir / xlsx.name).with_suffix('.pdf')

  sys.path.append(str(root.resolve()))
  from excel import Excel

  with Excel() as excel:
    excel.open(str(xlsx.resolve()))
    excel.to_pdf(str(pdf.resolve()))
  
  return pdf.resolve()

if __name__ == '__main__':
  excel2pef(xlsx_path)