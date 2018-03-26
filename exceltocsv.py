import os
import glob

import unicodecsv
import xlrd
import sys

def find_files_in_dir(dir, ext):
    os.chdir(dir)
    result = [i for i in glob.glob('*.{}'.format(ext))]
    return result

def process_file(inputfile):
    xl = xlrd.open_workbook(inputfile, encoding_override="cp1251")
    csv_filename = inputfile + '.csv'
    fh = open(csv_filename, "wb")
    csv_out = unicodecsv.writer(fh, delimiter=';', encoding='utf-8')
    sheet = xl.sheet_by_index(0)
    for row_number in range(sheet.nrows):
        csv_out.writerow(sheet.row_values(row_number))
    fh.close()
    xl.release_resources()
    # fh.release_resources()
    # csv_out.release_resources()
    del xl
    # del fh
    # del csv_out


if __name__ == "__main__":
    path = ""
    res = []
    if len(sys.argv) > 1:
        path = sys.argv[1]

    if path == "":
        res = find_files_in_dir(os.getcwd(), 'xls')
    else:
        if os.path.isfile(path):
            process_file(path)
        elif os.path.isdir(path):
            res = find_files_in_dir(path, 'xls')

    if len(res) > 0:
        for i in res:
            process_file(i)
