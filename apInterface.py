from tkinter import Tk
from invoicenet.gui.apInterface import apInterface
import pytesseract

pytesseract.pytesseract.tesseract_cmd = r'C:\Users\WalkerGA\AppData\Local\Tesseract-OCR\tesseract.exe'
#pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def main():
    root = Tk()
    apInterface(root)
    root.mainloop()

if __name__ == '__main__':
    main()