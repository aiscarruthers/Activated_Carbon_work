import sys
import fitz
result = fitz.open()

for pdf in sys.argv[1:]:
    with fitz.open(pdf) as mfile:
        result.insert_pdf(mfile)
    
result.save("merged.pdf")