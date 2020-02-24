import os 
from pdf2image import convert_from_path 
from pathlib import Path 
root = Path(__file__).parent.resolve()
poppler_bin = (root / 'poppler/poppler-0.68.0/bin').resolve()
os.environ['PATH'] += os.pathsep + str(poppler_bin.resolve())

def pdf2tiff(pdf, tiff_dir = None):
  
  pdf = Path(pdf)
  
  tiff_dir = pdf.parent / 'tiff' if not tiff_dir else Path(tiff_dir) 
  tiff_dir.mkdir(parents=True, exist_ok=True)

  dpi = 150
  pages = convert_from_path(str(pdf), dpi)
  max = len(pages)
  keta = len(str(max))

  for i, page in enumerate(pages):
    name = '{} ({}-{})'.format(
      pdf.stem, 
      str(i + 1).zfill(keta), 
      str(max).zfill(keta)
    )

    filename = (tiff_dir / name).with_suffix('.tiff')
    page.save(str(filename), 'TIFF')
  
