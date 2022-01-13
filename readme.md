# Typeconverter - A simple way to bulk-convert stuff to pdf

This repository was made pretty much out of a "Fine, I'll do it myself" moment. To bulk convert ppt/pptx at least, I found that I'd have to pay a premium for more than 10 files.

Well, guess I'll just do it myself. Right now we have:

1. PPT/PPTX to PDF
2. JPG/JPEG/PNG to PDF

Dependencies:

**Comtypes**

`pip install comtypes --user`

**FPDF**

`pip install fpdf2 --user`

**PIL**

`pip install Pillow --user`

**In case you get issues running this code**

`pip3 install pywinauto --user`

I might expand on this functionality a bit, there's quite a lot of room to improve on this. Contributors welcome!

**Usage instructions**

**PPT/PPTX to PDF**

Put the python script in the folder where you have the documents, and just run it. No frills.

**JPG/JPEG/PNG to PDF**

`python main.py --jpgpdf <outputfilename>.pdf`
