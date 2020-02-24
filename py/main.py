from pathlib import Path
import sys 

root = Path(__file__).parent.resolve() 
sys.path.append(str(root))

from excel2pdf import excel2pdf
from pdf2tiff import pdf2tiff

xlsx = (root / 'excel' / '011_納品書_タテ型.xlsx')
tiff_dir = (root / 'tiff')

def main():
  pdf = excel2pdf(str(xlsx.resolve()))
  if pdf.exists():
    pdf2tiff(str(pdf.resolve()), str(tiff_dir.resolve()))

if __name__ == '__main__':
  main()
