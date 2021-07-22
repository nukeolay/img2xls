import os
import math
import xlsxwriter
import argparse
import numpy as np
import sys
from PIL import Image
from flask import Flask, request, redirect, url_for, render_template, send_from_directory
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/uploads/'
DOWNLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/downloads/'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'bmp'}

app = Flask(__name__, static_url_path="/static")
DIR_PATH = os.path.dirname(os.path.realpath(__file__))
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
# limit upload size upto 8mb
app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            print('No file attached in request')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            print('No file selected')
            return redirect(request.url)
        try:
            width = int(request.form["width"])
            colors = int(request.form["colors"])
        except ValueError:
            print('Width and colors must be integer')
            return redirect(request.url)
        if width<2 or width>1024 or colors<2 or colors>256:
            print('Width must be within range from 2 to 1024 and colors must be from 2 to 256')
            return redirect(request.url)

        if file and allowed_file(file.filename):
           filename = secure_filename(file.filename)
           file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
           img2xls(os.path.join(app.config['UPLOAD_FOLDER'], filename), filename, width, colors)
           return redirect(url_for('uploaded_file', filename=filename + '.xlsx'))
    else:
        return render_template('index.html')

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

def rgbToHex(rgb):
    return '#%02x%02x%02x' % rgb

def img2xls(inputFilenameWithPath, filename, width, colorNumber):

    outputImageFilename = filename + '.png'
    outputImageFilenameWithPath = os.path.join(app.config['DOWNLOAD_FOLDER'], outputImageFilename)
    outputXlsxFilename = filename + '.xlsx'
    outputXlsxFilenameWithPath = os.path.join(app.config['DOWNLOAD_FOLDER'], outputXlsxFilename)

    inputFile = Image.open(inputFilenameWithPath)

    ratio = inputFile.size[1]/inputFile.size[0]
    height = math.floor(width * ratio)
    resized_image = inputFile.resize((width,height))
    resized_image = resized_image.convert(mode='P', palette=Image.ADAPTIVE, colors=colorNumber)
    palette = np.array(resized_image.getpalette(),dtype=np.uint8).reshape((256,3))

    workbook = xlsxwriter.Workbook(outputXlsxFilenameWithPath)
    worksheetColored = workbook.add_worksheet('Colored')
    worksheetUncolored = workbook.add_worksheet('Uncolored')
    cell_format = workbook.add_format()
    cell_formatUncolored = workbook.add_format()


    cellSize = 30 # ширина и высота ячейки
    pix = resized_image.load()
    maxWidth = resized_image.size[0]
    maxHeight = resized_image.size[1]

    worksheetColored.set_column_pixels(0, 0, 70) # устанавливаем ширину первой колонки
    worksheetUncolored.set_column_pixels(0, 0, 70)

    cell_format.set_bg_color('white')
    cell_format.set_text_wrap()
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    worksheetColored.write(0, 0, filename, cell_format) # пишем название файла в первой строке файла
    worksheetUncolored.write(0, 0, filename, cell_format) # пишем название файла в первой строке файла

    cell_formatUncolored.set_bg_color('white')
    cell_formatUncolored.set_text_wrap()
    cell_formatUncolored.set_align('center')
    cell_formatUncolored.set_align('vcenter')
    cell_formatUncolored.set_border(4)
    cell_formatUncolored.set_border_color('#808080')

    for row in range (0, maxHeight):
        worksheetColored.set_row_pixels(row,cellSize)
        worksheetUncolored.set_row_pixels(row,cellSize)
        for col in range(0, maxWidth):
            cellNum=pix[col,row]
            cellColor = rgbToHex((palette[cellNum][0],palette[cellNum][1],palette[cellNum][2]))
            cell_format = workbook.add_format()
            cell_format.set_bg_color(str(cellColor))
            cell_format.set_align('center')
            cell_format.set_align('vcenter')
            worksheetColored.set_column_pixels(col + 2, col + 2, cellSize)
            worksheetColored.write(row + 2, col + 2, '', cell_format)
            worksheetUncolored.set_column_pixels(col + 2, col + 2, cellSize)
            worksheetUncolored.write(row + 2, col + 2, str(cellNum), cell_formatUncolored)

    
    
    cell_formatTab = workbook.add_format()
    cell_formatTab.set_text_wrap()
    cell_formatTab.set_align('center')
    cell_formatTab.set_align('vcenter')
    worksheetColored.set_row_pixels(maxHeight + 3, 50) # устанавливаем высоту строки
    worksheetColored.write(maxHeight + 3, 0, 'Color number', cell_formatTab) # пишем заголовок таблицы с цветами
    worksheetColored.write(maxHeight + 3, 1, 'Color HEX code', cell_formatTab) # пишем заголовок таблицы с цветами
    worksheetUncolored.set_row_pixels(maxHeight + 3, 50) # устанавливаем высоту строки
    worksheetUncolored.write(maxHeight + 3, 0, 'Color number', cell_formatTab) # пишем заголовок таблицы с цветами
    worksheetUncolored.write(maxHeight + 3, 1, 'Color HEX code', cell_formatTab) # пишем заголовок таблицы с цветами
    
    for i in range(0, colorNumber) :
        row = i + maxHeight + 4 # с какой строки начинать запись цветов
        cellColor = rgbToHex((palette[i][0],palette[i][1],palette[i][2]))
        cell_formatTab = workbook.add_format()
        cell_formatTab.set_align('center')
        cell_formatTab.set_align('vcenter')
        cell_formatTab.set_bg_color(str(cellColor))
        worksheetColored.write(row, 0, i, cell_formatTab)
        worksheetColored.write(row, 1, cellColor, cell_formatTab)
        worksheetUncolored.write(row, 0, i, cell_formatTab)
        worksheetUncolored.write(row, 1, cellColor, cell_formatTab)

    resized_image.save(outputImageFilenameWithPath)
    resized_image.close()
    inputFile.close()
    workbook.close()