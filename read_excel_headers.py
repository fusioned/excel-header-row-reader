#!/usr/bin/python3
"""
Excel workbook header line bulk reader
======================================

Reads and outputs the header row for the given Excel files
Optionally append the first content row after the header for additional inspection

"""

import sys
import getopt
from openpyxl import load_workbook

class CsvRender:
    """Output to a CSV format"""

    def __init__(self, num_cols_pad=20, append_first_body_row=False):
        """
        num_cols_pad indicate the number of columns minimum before the body data is added on
        e.g. when num_cols_pad = 3 and append_first_body_row = True
            ,,,extra,body_content
            a,b,,extra,content
            filename,,,additional

        """
        self.num_cols = num_cols_pad
        self.append_body = append_first_body_row
        self.cols = 0

    def start(self):
        pass

    def filename(self, filename):
        self.cols = 0
        self._cell(filename)

    def sheet_name(self, name):
        self._cell(name)

    def header_cell(self, value, cell, index):
        self._cell(value if value else '')

    def header_row(self, row):
        pass

    def body_row(self, row):
        # pad this out
        if self.append_body:
            print(',' * (self.num_cols - self.cols), end='')

    def body_cell(self, value, cell, index):
        if self.append_body:
            self._cell(value if value else '')

    def done(self):
        print('')

    def _cell(self, value):
        if self.cols > 0:
            print(',', end='')

        print('"' + value + '"', end='')
        self.cols += 1


class ClassicRender:
    """Output data indented"""
    def start(self):
        pass

    def filename(self, filename):
        print('[' + filename + ']')

    def sheet_name(self, name):
        print(name + ':')

    def header_cell(self, value, cell, index):
        if value:
            print('  ' + value)

    def header_row(self, row):
        pass

    def body_row(self, row):
        pass

    def body_cell(self, value, cell, index):
        if (index > 0):
            print(',', end='')
        print(value, end='')

    def done(self):
        print('')


def main(argv):
    if len(argv) <= 1:
        print('Usage: ' + argv[0] + ' <workbook>')
        quit()

    renderer = CsvRender(append_first_body_row = True)

    for filename in argv[1:]:
        process(filename, renderer)


def process(filename, renderer):
    renderer.filename(filename)
    wb = load_workbook(filename)
    sheets = wb.get_sheet_names()

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        renderer.sheet_name(sheet_name)

        header = sheet[1]
        renderer.header_row(header)
        for index, cell in enumerate(header):
            renderer.header_cell(cell.value, cell, index)

        row = sheet[2]
        renderer.body_row(row)
        for index, cell in enumerate(row):
            renderer.body_cell(cell.value, cell, index)

    renderer.done()


if __name__ == '__main__':
    main(sys.argv)
