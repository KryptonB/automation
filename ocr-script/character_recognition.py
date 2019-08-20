try:
    from PIL import Image
except ImportError:
    import Image
import pytesseract
import os

# set the path to the tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\sratnappuli\AppData\Local\Tesseract-OCR\tesseract.exe'

# character recognition function
def oc_recog(filename):

    # convert the image to string
    text = pytesseract.image_to_string(Image.open(filename))
    
    # create a new txt file and write the above string to it
    f = open('output.txt', 'w')
    f.write(text)
    f.close()
    
    return text

# call the character recognition function 
print(oc_recog('./tt.png'))