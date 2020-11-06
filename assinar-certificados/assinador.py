import fitz
from os import sep
from glob import glob
from os import path
from os import makedirs

#https://stackabuse.com/working-with-pdfs-in-python-adding-images-and-watermarks/

output_folder = "output" + sep
input_folder = "input" + sep
assinatura_folder = "assinatura" + sep 

if not path.exists(output_folder):
    makedirs(output_folder)

if not path.exists(input_folder):
    print ("Não existe input")
    exit(-1)


if not path.exists(assinatura_folder):
    print ("Não existe assinatura")
    exit(-2)

jessica = assinatura_folder + "image.png"
rodrigo = assinatura_folder + "image.png"
marcelo = assinatura_folder + "image.png"

# define the posdition (upper-right corner)
image_jessica = fitz.Rect(30,320,250,632)
image_rodrigo = fitz.Rect(240,320,470,632)
image_marcelo = fitz.Rect(430,320,660,632)

files = glob(input_folder + "*.pdf")
output_file = "example2.pdf"
for input_file in files: 
    # retrieve the first page of the PDF
    file_handle = fitz.open(input_file)
    first_page = file_handle[0]
    output_file = output_folder + path.split(input_file)[1]
    # add the image
    first_page.insertImage(image_jessica, filename = jessica)
    first_page.insertImage(image_rodrigo, filename = rodrigo)
    first_page.insertImage(image_marcelo, filename = marcelo)

    file_handle.save(output_file, deflate=True)
