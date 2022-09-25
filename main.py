import openpyxl as openpyxl
from PIL import Image
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import webbrowser
import os
import cv2

if __name__ == '__main__':

    wb = openpyxl.Workbook()

    vidcap = cv2.VideoCapture('dolphin21-preview_COkhxs4Y.mp4')
    success, image = vidcap.read()
    count = 0
    while success:
        cv2.imwrite("Images/%d.jpg" % count, image)  # save frame as JPEG file
        success, image = vidcap.read()
        print('Read a new frame: ', success)
        count += 1

    images = os.listdir('Images')

    #need to get images in numeric order

    for i in range(len(images)):
        print(i)

        wb.create_sheet(str(i))
        wb.active = wb[str(i)]
        ws = wb.active
        ws.title = str(i)

        with Image.open('Images/' + str(i) + '.jpg') as inputImage:
            # inputImage.show()
            rgbInputImage = inputImage.convert('RGB')
            width = inputImage.width
            height = inputImage.height

        for hPixel in range(1, height):
            shift = hPixel - 1 if hPixel > 1 else hPixel
            index = len(ws['A']) - shift
            for wPixel in range(1, width):
                r, g, b = rgbInputImage.getpixel((wPixel, hPixel))
                columnLetter = get_column_letter(wPixel)

                rVal = 'FF%02x%02x%02x' % (r, 0, 0)
                gVal = 'FF%02x%02x%02x' % (0, g, 0)
                bVal = 'FF%02x%02x%02x' % (0, 0, b)

                # ws[columnLetter + str(index + hPixel)] = r
                ws[columnLetter + str(index + hPixel)].fill = PatternFill(start_color=rVal, end_color=rVal,
                                                                          fill_type="solid")
                # ws[columnLetter + str(index + hPixel + 1)] = g
                ws[columnLetter + str(index + hPixel + 1)].fill = PatternFill(start_color=gVal, end_color=gVal,
                                                                              fill_type="solid")
                # ws[columnLetter + str(index + hPixel + 2)] = b
                ws[columnLetter + str(index + hPixel + 2)].fill = PatternFill(start_color=bVal, end_color=bVal,
                                                                              fill_type="solid")
        ws.sheet_view.zoomScale = 5
        ws.sheet_view.showGridLines = False

    wb.remove(wb['Sheet'])
    wb.save(filename='sample_book.xlsx')
    webbrowser.open('sample_book.xlsx')
    # sleep(1)
    # os.system("taskkill /im EXCEL.EXE /f")
