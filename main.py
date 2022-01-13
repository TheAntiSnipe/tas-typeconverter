import comtypes.client
import os


def PPTtoPDF(inputFileName, outputFileName, formatType=32):
    print(inputFileName)
    pathdata = os.path.abspath(__file__)
    dirdata = os.path.dirname(pathdata)
    filedata = os.path.join(dirdata, inputFileName)
    outputfiledata = os.path.join(dirdata, outputFileName)

    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(filedata)
    deck.SaveAs(outputfiledata, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()


if(__name__ == '__main__'):
    filename_list = [i for i in os.listdir() if i[-4:] ==
                     'pptx' or i[-3:] == 'ppt']
    replaced_list = [('pdfdoc_'+i).split('.')[0] for i in filename_list]
    for filename_number in range(len(filename_list)):
        PPTtoPDF(filename_list[filename_number],
                 replaced_list[filename_number])
