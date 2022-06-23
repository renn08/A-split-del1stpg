import pygsheets
import pandas as pd
import os
import PyPDF2


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

    # filename = ''
    # for file in os.listdir('./original'):
    #     if not filename:
    #         filename = file
    #     print(file)

    # # creating a pdf file object
    # pdfFileObj = open('./original/' + filename, 'rb')
    #
    # # creating a pdf reader object
    # pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    #
    # # printing number of pages in pdf file
    # print(pdfReader.numPages)
    #
    # # creating a page object
    # pageObj = pdfReader.getPage(0)
    #
    # # extracting text from page
    # print(pageObj.extractText())
    #
    # writer = PyPDF2.PdfFileWriter()
    #
    # start = 16
    #
    # end = 16
    #
    # output_filename = ''
    # while start <= end:
    #     writer.addPage(pdfReader.getPage(start))
    #     start += 1
    #     output_filename = "example.pdf"
    #
    # with open(output_filename, 'wb') as out:
    #     writer.write(out)
    #
    #
    # # closing the pdf file object
    # pdfFileObj.close()


def main():
    # create_map()
    write_filename_to_sheet()
    # rename_files()


if __name__ == '__main__':
    main()
