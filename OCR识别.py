import pytesseract
from PIL import Image

# pytesseract.pytesseract.tesseract_cmd = 'D://ocr/Tesseract-OCR/tesseract.exe'
# text = pytesseract.image_to_string(Image.open('D:\\Users\\Administrator\\Desktop\\需要用到的文件\\OCR识别\\44.png'), lang='chi_sim')
# print(text)



text = pytesseract.image_to_string(Image.open(r'D:\\Users\\Administrator\\Desktop\\需要用到的文件\\OCR识别\\44.png'))
print(text)