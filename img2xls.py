import math
import xlsxwriter
import argparse
import numpy as np
import sys
from PIL import Image

def rgbToHex(rgb):
    return '#%02x%02x%02x' % rgb


def img2xls(inputFileName, outputFileName, width, colorNumber):
    outputImageFileName = outputFileName + '.png'
    outputXlsxFileName = outputFileName + '.xlsx'

    inputFile = Image.open(inputFileName)
    ratio = inputFile.size[1]/inputFile.size[0]
    height = math.floor(width * ratio)
    resized_image = inputFile.resize((width,height))
    resized_image = resized_image.convert(mode='P', palette=Image.ADAPTIVE, colors=colorNumber)
    palette = np.array(resized_image.getpalette(),dtype=np.uint8).reshape((256,3))

    workbook = xlsxwriter.Workbook(outputXlsxFileName)
    worksheet = workbook.add_worksheet()
    cell_format = workbook.add_format()

    pix = resized_image.load()
    maxWidth = resized_image.size[0]
    maxHeight = resized_image.size[1]
    for row in range (1, maxHeight):
        worksheet.set_row_pixels(row,7)
        stringRow = ''
        for col in range(1, maxWidth):
            cellNum=pix[col,row]
            cellColor = rgbToHex((palette[cellNum][0],palette[cellNum][1],palette[cellNum][2]))
            cell_format = workbook.add_format()
            cell_format.set_bg_color(str(cellColor))
            worksheet.set_column_pixels(col,col,7)
            worksheet.write(row, col, str(cellNum), cell_format)
            stringRow = stringRow + str(cellNum)
        print(stringRow)

    resized_image.save(outputImageFileName)
    resized_image.close()
    inputFile.close()
    workbook.close()

    print('------------')
    print('PALETTE:')
    for i in range(0, colorNumber) :
        print(str(i) + ': ' + str(palette[i]) + ' = ' + str(rgbToHex((palette[i][0],palette[i][1],palette[i][2]))))
    print('------------')    
    print('SUCCESS!')
    print('- image size: '+ str(maxWidth) + 'x' + str(maxHeight))
    print('- color number: '+ str(colorNumber))
    print('- image file name: ' + outputImageFileName)
    print('- xlsx file name: ' + outputXlsxFileName)


    


my_parser = argparse.ArgumentParser(description='img2xls CLI, v.0.7')
my_parser.add_argument('inputFileName', metavar='inputFileName', type=str, help='name of the file to be converted')
my_parser.add_argument('-o', metavar='output file name', type=str, help='name of the converted file', default='')
my_parser.add_argument('-w', metavar='width', type=int, help='width of the output file (default: 15 pixels)', default=15)
my_parser.add_argument('-c', metavar='colors', type=int, help='color number of the output image (default: 8 colors)', default=8)

args = my_parser.parse_args()

inputFileName = args.inputFileName
if args.o == '':
    outputFileName = 'new ' + inputFileName
else:
    outputFileName = args.o
width = args.w
colorNumber = args.c

img2xls(inputFileName, outputFileName, width, colorNumber)
