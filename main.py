import shutil
import openpyxl as openpyxl
from PIL import Image
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import os
import cv2
import moviepy.editor as mp


def splitVideo(video):
    vidcap = cv2.VideoCapture(video)
    success, image = vidcap.read()
    count = 0
    while success:
        cv2.imwrite("Images/%d.jpg" % count, image)
        success, image = vidcap.read()
        count += 1


def deleteFilesInDirectory(directory):
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))


if __name__ == '__main__':

    wb = openpyxl.Workbook()
    dontCont = True
    lameMode = False

    while dontCont:
        val = input("1: use new video \n2: use existing photos in Image folder\n->  ")
        if val == '1' or val == '2':
            dontCont = False
        else:
            print('Sorry, please enter a valid input')

    # sample video: 'dolphin21-preview_COkhxs4Y.mp4'

    if val == '1':
        video = input("Please enter video name:\n->  ")
        resizeYN = input("Would you like to resize the video(recommended size is 144p)?:\n(y/n)->  ")
        if resizeYN == "y" or resizeYN == "Y":
            # do the resize
            resizeW = ""
            resizeH = ""
            while not resizeW.isdigit():
                resizeW = input("Please enter the width:\n->  ")
            while not resizeH.isdigit():
                resizeH = input("Please enter the height:\n->  ")
            clip = mp.VideoFileClip(video)
            clip_resized = clip.resize(height=int(resizeH))
            clip_resized = clip_resized.resize(width=int(resizeW))
            clip_resized.write_videofile("resized-" + video)
            deleteFilesInDirectory('Images')
            splitVideo("resized-" + video)
        else:
            splitVideo(video)

    images = os.listdir('Images')

    for i in range(len(images)):
        print(i)

        wb.create_sheet(str(i))
        wb.active = wb[str(i)]
        ws = wb.active
        ws.title = str(i)

        with Image.open('Images/' + str(i) + '.jpg') as inputImage:

            rgbInputImage = inputImage.convert('RGB')
            width = inputImage.width
            height = inputImage.height

        for hPixel in range(1, height):
            shift = hPixel - 1 if hPixel > 1 else hPixel
            index = len(ws['A']) - shift
            for wPixel in range(1, width):
                r, g, b = rgbInputImage.getpixel((wPixel, hPixel))
                columnLetter = get_column_letter(wPixel)

                if lameMode:
                    rgbVal = 'FF%02x%02x%02x' % (r, g, b)
                    ws[columnLetter + str(index + hPixel)].fill = PatternFill(start_color=rgbVal, end_color=rgbVal,
                                                                              fill_type="solid")
                else:
                    rVal = 'FF%02x%02x%02x' % (r, 0, 0)
                    gVal = 'FF%02x%02x%02x' % (0, g, 0)
                    bVal = 'FF%02x%02x%02x' % (0, 0, b)

                    ws[columnLetter + str(index + hPixel)].fill = PatternFill(start_color=rVal, end_color=rVal,
                                                                              fill_type="solid")
                    ws[columnLetter + str(index + hPixel + 1)].fill = PatternFill(start_color=gVal, end_color=gVal,
                                                                                  fill_type="solid")
                    ws[columnLetter + str(index + hPixel + 2)].fill = PatternFill(start_color=bVal, end_color=bVal,
                                                                                  fill_type="solid")
        ws.sheet_view.zoomScale = 5
        ws.sheet_view.showGridLines = False

        #add a reduced ram usage mode
        #wb.save(filename='sample_book.xlsx')

    wb.remove(wb['Sheet'])
    wb.save(filename='sample_book.xlsx')
    # webbrowser.open('sample_book.xlsx')
