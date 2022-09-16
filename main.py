import pygsheets
import pandas as pd
import os
import PyPDF2
from pathlib import Path


def write_to_sheet():
    # authorization
    gc = pygsheets.authorize(service_file='./examtranscript-bf5d3d2ff627.json')

    # Create empty dataframe
    df = pd.DataFrame()

    # Create a column
    df['name'] = ['John', 'Steve', 'Sarah']

    # open the google spreadsheet with its name
    sh = gc.open('finalexam-A-transcription')

    # select the first sheet
    wks = sh[0]

    # update the first sheet with df, starting at cell B2.
    wks.set_dataframe(df, (3, 3))


def create_map():
    idx = 1
    with open('map.txt', 'w') as fw:
        for pdf in os.listdir('./original'):
            fw.writelines(pdf + ' ' + str(idx) + '\n')
            idx += 1


def rename_files():
    with open('map.txt', 'r') as fr:
        lines = fr.readlines()

    for line in lines:
        line_striped = line.strip()
        p1, p2, p3, p4, p5, index = line_striped.split(' ')
        filename = p1 + ' ' + p2 + ' ' + p3 + ' ' + p4 + ' ' + p5
        os.rename('./renamed/' + filename, './renamed/' + str(index) + '.pdf')


def write_filename_to_sheet():
    # authorization
    gc = pygsheets.authorize(service_file='./examtranscript-bf5d3d2ff627.json')

    with open('map.txt', 'r') as fr:
        lines = fr.readlines()

    # Create empty dataframe
    df = pd.DataFrame()
    filename_list = []

    for line in lines:
        line_striped = line.strip()
        p1, p2, p3, p4, p5, index = line_striped.split(' ')
        filename = p1 + ' ' + p2 + ' ' + p3 + ' ' + p4 + ' ' + p5
        filename_list.append(filename)
        # df['FileName'].append(pd.Series([filename]))

    df['FileName'] = filename_list

    # open the google spreadsheet with its name
    sh = gc.open('finalexam-A-transcription')

    # select the first sheet
    wks = sh[0]

    # update the first sheet with df, starting at cell A1.
    wks.set_dataframe(df, (1, 1))


def extract_page_as_new_pdf(src, start, end, out):
    # creating a pdf file object
    pdfFileObj = open(src, 'rb')
    # creating a pdf reader object
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    # creating a pdf writer
    writer = PyPDF2.PdfFileWriter()

    while start <= end:
        writer.addPage(pdfReader.getPage(start))
        start += 1

    with open(out, 'wb') as out:
        writer.write(out)

    pdfFileObj.close()


def extract_q22():
    index = 1
    while index != 210:
        extract_page_as_new_pdf(src='./renamed/' + str(index) + '.pdf',
                                start=16, end=16, out='./q22/' + str(index) +
                                                      '.pdf')
        index += 1


def extract_q8():
    if not os.path.isdir('q8'):
        os.mkdir('q8')
    index = 1
    while index != 210:
        extract_page_as_new_pdf(src='./renamed/' + str(index) + '.pdf',
                                start=4, end=4, out='./q8/' + str(index) +
                                                      '.pdf')
        index += 1


def combine(path):
    merger = PyPDF2.PdfFileMerger()
    for i in range(209):
        name = str(i + 1) + '.pdf'
        filepath = str(Path(path) / name)
        merger.append(PyPDF2.PdfFileReader(open(filepath, 'rb')))
    merger.write(str(Path(path) / 'output.pdf'))


def main():
    # create_map()
    # write_filename_to_sheet()
    # rename_files()
    # extract_q22()
    # extract_q8()
    combine('q8')
    pass


if __name__ == '__main__':
    main()
