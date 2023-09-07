import os
import comtypes.client
from pdf2image import convert_from_path




def init_powerpoint():
   powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
   powerpoint.Visible = 1
   return powerpoint


def ppt_to_pdf(powerpoint,inputFileName, outputFileName):
    if outputFileName[-3:] != 'pdf':
        outputFileName += ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, 32) # formatType = 32 for pdf
    deck.Close()

def convert_files_in_folder(powerpoint, folder):
    files = os.listdir(folder)
    for file in files:
        if file[-3:] == "ppt":
            file_path = f'{folder}\\{file}'
            ppt_to_pdf(file_path, file_path)



def pdf_to_img(pdf_file,output_folder):
    images = convert_from_path(pdf_file)
    for i, image in enumerate(images):
        fname = output_folder+'image'+str(i)+'.png'
        image.save(fname, "PNG")




if __name__ == "__main__":
    powerpoint = init_powerpoint()
    
    inputFileName = "C:\\PPTVoiceOver\\pptData\\example_1.pptx"   #e.g "C:\\Users\\User\\Desktop\\file.ppt"
    outputFileName = "C:\\PPTVoiceOver\\temp_pdf\\example_1.pdf" #e.g "C:\\Users\\User\\Desktop\\output.pdf"
    imagesOutputpath = "C:\\PPTVoiceOver\\Images\\example_1\\"
    ppt_to_pdf(powerpoint,inputFileName, outputFileName)
       # pdf_file = "your pdf file"  # e.g "C:\\Users\\User\\Desktop\\output.pdf"
    #pdf_to_img(outputFileName,imagesOutputpath)