import comtypes.client
import fpdf
from PIL import Image
import os
import sys


def JPGtoPDF(imagelist, filepath):
    pathdata = os.path.abspath(__file__)
    dirdata = os.path.dirname(pathdata)
    outputfiledata = os.path.join(dirdata, filepath)
    pdf = fpdf.FPDF()
    # imagelist is the list with all image filenames
    for image in imagelist:
        cover = Image.open(image)
        width, height = cover.size
        width, height = float(width * 0.264583), float(height * 0.264583)
        pdf.add_page(format=(width, height))
        pdf.image(image, 0, 0, width, height)
    pdf.output(outputfiledata, "F")


def PPTtoPDF(inputFileName, outputFileName, formatType=32):
    print(inputFileName)
    pathdata = os.path.abspath(__file__)
    dirdata = os.path.dirname(pathdata)
    filedata = os.path.join(dirdata, inputFileName)
    outputfiledata = os.path.join(dirdata, outputFileName)

    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if outputFileName[-3:] != "pdf":
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(filedata)
    deck.SaveAs(outputfiledata, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == "--jpgpdf" and len(sys.argv) > 2:
            filename_list = [
                i
                for i in os.listdir()
                if i[-4:] == "jpeg" or i[-3:] == "jpg" or i[-3:] == "png"
            ]
            outputname = sys.argv[2]
            JPGtoPDF(filename_list, outputname)
        elif sys.argv[1] == "--help":
            print(
                """\nList of valid arguments:
    --jpgpdf : Converts jpg/jpeg/png to pdf, requires an output file name to follow.\n"""
            )
        elif sys.argv[1] == "--jpgpdf":
            print(
                """
You need to specify a filepath as well!

Correct way to call the command:
python main.py --jpgpdf <yourfilename>.pdf\n"""
            )
        else:
            print(
                """
Argument not recognized!

List of valid arguments:
    --jpgpdf : Converts jpg/jpeg/png to pdf, requires an output file name.\n"""
            )

    else:
        filename_list = [i for i in os.listdir() if i[-4:] == "pptx" or i[-3:] == "ppt"]
        replaced_list = [("pdfdoc_" + i).split(".")[0] for i in filename_list]
        for filename_number in range(len(filename_list)):
            PPTtoPDF(filename_list[filename_number], replaced_list[filename_number])
