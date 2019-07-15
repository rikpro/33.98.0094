#!/usr/bin/python2.7
"""
Provides xlsx file output of P52 JSON file data
"""
__author__ = "Rik Provoost"
__company__ = "Televic Rail"

import os
import sys
import json
import ast
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.drawing.image import Image


def style_range(cell_range, font=None, fill=None, \
                border=None, alignment=None, merge=False):
    """
    Apply styles to a range of cells as if they were a single cell.
    """
    for cells in cell_range:
        for cell in cells:
            if font:
                cell.font = font
            if fill:
                cell.fill = fill
            if border:
                cell.border = border
            if alignment:
                cell.alignment = alignment
#     if merge == True:
#         ws.merge_cells(cell_range)


class TestSpecsXlsx(object):
    """
    Class TestSpecsXlsx
    """
    productionfiles_folder = "S:\\R&D\\Ontwikkelingen\\%s\\G.productiedossier\\"

    doc_type = {'Name': 'P52', 'Description': 'ROUTINE TEST SPECIFICATIONS'}
    title_cell = ('E', '2')
    column_titles = {'B': 'STEP', 'C': 'TITLE', 'D': 'DESCRIPTION', 'E': 'EXPECTED RESULT(S)'}

    START_LINE = 6  # Line number where to start the Test Cases and Test Steps

    try:
        LOCAL_LOCATION = os.path.dirname(os.path.realpath(__file__)) + "\\"
    except:
        LOCAL_LOCATION = os.path.dirname(os.path.realpath(sys.argv[0])) + "\\"

    PICS_FOLDER = LOCAL_LOCATION + "pics\\"
    TLV_LOGO = PICS_FOLDER + "Televic-rail.jpg"

    def __init__(self, jsonfile=None, outputfolder="C:\\temp"):
        """
        Init Method of class.
        """
        self.line = self.START_LINE
        self.jsondata = self._readJson(jsonfile=jsonfile)
        self.outputfolder = outputfolder
        # get file next version
        self.version = self._getFileNextVersion(partnumber=self.jsondata['Partnumber'], filetype=self.doc_type['Name'])
        print "Next file version: ", self.version

        filename_xlsx = self.jsondata['Partnumber'] + "_" + self.doc_type['Name'] + "_" + self.version + ".xlsx"
        self.filename = os.path.join(self.outputfolder, filename_xlsx)

        self._open()

    def _open(self):
        """
        open xls files
        """
        self.xlwb = Workbook()
        self.xlws = self.xlwb.worksheets[0]  # get first worksheet

    def close(self):
        """
        Close and Save file.
        """
        print "Closing ", self.filename
        try:
            self.xlwb.save(self.filename)
        except Exception as e:
            print("ERROR: ", e)

    def _getTestCaseDescription(self, testcase='FT'):
        """
        Get Test Cases Description
        :return: (str) Test Case Description
        """
        testtypes_lst = ast.literal_eval(self.jsondata['TestCaseList'])
        for testtype in testtypes_lst:
            if testtype['name'] == testcase:
                return testtype['description']

    def _getFileNextVersion(self, partnumber="33.92.9999", filetype="P52"):
        """
        Get lastest version the file with filetype on the fileserver.

        :return: (str) next file version
        """
        folder = self.productionfiles_folder % partnumber
        filename_begin_str = partnumber + "_" + filetype
        old_fversion = 0
        new_fversion = 0
        try:
            for afile in os.listdir(folder):
                if afile.startswith(filename_begin_str):
                    point_idx = afile.rfind(".")
                    version = afile[point_idx - 4:point_idx]
                    try:
                        new_fversion = float(version) + 0.01
                    except ValueError:
                        print "Not a float"
                    if new_fversion > old_fversion:
                        old_fversion = new_fversion
        except Exception:  # pylint: disable=broad-except
            pass
            # print "Folder does not exist"

        if new_fversion == 0:
            return "1.01"
        else:
            return str(new_fversion)

    def _readJson(self, jsonfile=''):
        """
        read the JSON file and convert to dict
        :param jsonfile: (str) JSON full filename
        :return: jsondata
        """
        print "Read JSON file:", jsonfile
        with open(jsonfile, 'r') as f:
            jsondata = json.load(f)
        # print "JSON data:", jsondata
        return jsondata

    def _writeTitle(self):
        """
        Write title block and page footer.
        """
        LOGO_CELL = 'A1'
        DOCUMENT_TYPE_CELL = 'D2'
        DOCUMENT_TITLE_CELL = 'D3'

        img = Image(TestSpecsXlsx.TLV_LOGO)
        self.xlws.add_image(img, LOGO_CELL)

        font_normal = Font(name='Arial', size=11)
        font_bold = Font(name='Arial', size=11, bold=True)

        alignment_left = Alignment(horizontal="left",
                                   vertical="top",
                                   wrap_text=False)

        alignment_right = Alignment(horizontal="right",
                                    vertical="top",
                                    wrap_text=False)

        # set data
        self.xlws[DOCUMENT_TYPE_CELL].value = self.doc_type['Description'] + ' ' + self.doc_type['Name']
        self.xlws[DOCUMENT_TITLE_CELL].value = self.jsondata['Partnumber'] + '   ' + self.jsondata['Description']

        style_range(self.xlws[DOCUMENT_TYPE_CELL: DOCUMENT_TYPE_CELL],
                    font=font_bold,
                    alignment=alignment_left,
                    merge=False)

        style_range(self.xlws[DOCUMENT_TITLE_CELL: DOCUMENT_TITLE_CELL],
                    font=font_normal,
                    alignment=alignment_left,
                    merge=False)

        # set data
        title = {self.title_cell[0] + str(int(self.title_cell[1]) + 0): 'Doc. Number: '
                                                                        + self.jsondata['Partnumber']
                                                                        + "_" + self.doc_type['Name'],
                 self.title_cell[0] + str(int(self.title_cell[1]) + 1): 'Version: ' + self.version,
                 self.title_cell[0] + str(int(self.title_cell[1]) + 2): 'Date: ' + datetime.now().strftime("%d-%m-%Y")}

        for cell, value in title.items():
            self.xlws[cell].value = value

        # set formatting
        style_range(self.xlws[self.title_cell[0] + str(int(self.title_cell[1]) + 0): self.title_cell[0] + str(
            int(self.title_cell[1]) + 3)],
                    font=font_normal,
                    alignment=alignment_right,
                    merge=False)

        style_range(self.xlws[chr(ord(self.title_cell[0]) + 1) + str(int(self.title_cell[1]) + 0): chr(
            ord(self.title_cell[0]) + 1) + str(int(self.title_cell[1]) + 3)],
                    font=font_normal,
                    alignment=alignment_left,
                    merge=False)

        # set header and footer
        self.xlws.oddFooter.left.text = "&[File]"
        self.xlws.oddFooter.left.font_size = 8
        self.xlws.oddFooter.left.font_name = "Calibri"

        self.xlws.oddFooter.right.text = "&[Page]/&N"
        self.xlws.oddFooter.right.font_size = 8
        self.xlws.oddFooter.right.font_name = "Calibri"

        self.xlws.evenFooter.left.text = "&[File]"
        self.xlws.evenFooter.left.font_size = 8
        self.xlws.evenFooter.left.font_name = "Calibri"

        self.xlws.evenFooter.right.text = "&[Page]/&N"
        self.xlws.evenFooter.right.font_size = 8
        self.xlws.evenFooter.right.font_name = "Calibri"

    def _writeTestCaseOverview(self):
        """
        Write TestCase overview list.
        """
        cell = 'B'

        font_title_bold = Font(name='Arial', size=16, bold=True)
        font_normal = Font(name='Arial', size=11, bold=False)

        alignment_left = Alignment(horizontal="left",
                                   vertical="top",
                                   wrap_text=False)

        self.xlws[cell + str(self.line)] = 'Test Cases:'
        style_range(self.xlws[cell + str(self.line):cell + str(self.line)],
                    font=font_title_bold,
                    alignment=alignment_left,
                    merge=False)
        self.line += 1

        # overview of TestCases
        for testtype in ast.literal_eval(self.jsondata['TestCaseList']):

            self.xlws[cell + str(self.line)] = '   ' + testtype['description'] + ' (' + testtype['name'] + ')'

            style_range(self.xlws[cell + str(self.line):cell + str(self.line)],
                        font=font_normal,
                        alignment=alignment_left,
                        merge=False)
            self.line += 1
        self.line += 2

    def _writeTestCaseTitles(self, line, testcase='FT'):
        """
        Method
        """
        cell = 'B'

        font_new = Font(name='Arial', size=16, bold=True)

        alignment_center = Alignment(horizontal="left",
                                     vertical="top",
                                     wrap_text=False)

        self.xlws[cell + str(line)] = self._getTestCaseDescription(testcase=testcase) + ' (' + testcase + ')'

        style_range(self.xlws[cell + str(line):cell + str(line)],
                    font=font_new,
                    alignment=alignment_center,
                    merge=False)

    def _writeColumnTitles(self, line):
        """
        Write Column Titles.
        """
        border_title = Border(left=Side(border_style='thin'),
                              right=Side(border_style='thin'),
                              top=Side(border_style='double'),
                              bottom=Side(border_style='thin'))

        font_normal = Font(name='Arial', size=10, bold=True)

        alignment_center = Alignment(horizontal="center",
                                     vertical="top",
                                     wrap_text=True)

        # set data and formatting
        for cell, value in self.column_titles.items():
            self.xlws[cell + str(line)] = value

            style_range(self.xlws[cell + str(line): cell + str(line)],
                        border=border_title,
                        font=font_normal,
                        alignment=alignment_center,
                        merge=False)

    def write(self):
        """
        Write complete xlsx file
        """
        self._setupPage()
        self._writeTitle()
        self._writeTestCaseOverview()
        self._writeTestSteps()
        self.close()

    def _setupPage(self):
        """
        Setup page
        """
        self.xlws.page_setup.paperSize = self.xlws.PAPERSIZE_A4
        self.xlws.page_setup.orientation = self.xlws.ORIENTATION_LANDSCAPE

        # Set cell width
        column_widths = {'A': 1, 'B': 6, 'C': 20,
                         'D': 65, 'E': 49, 'F': 10}

        for cell, width_value in column_widths.items():
            self.xlws.column_dimensions[cell].width = width_value

        # set page margins
        self.xlws.page_margins.left = 0.25
        self.xlws.page_margins.right = 0.25
        self.xlws.page_margins.top = 1
        self.xlws.page_margins.bottom = 1

    def _writeTestSteps(self):
        """
        Write all TestSteps for all TestCases defined in JSON file
        """
        cell_column = 'A'  # cell column to start from

        border_thin = Border(left=Side(border_style='thin'),
                             right=Side(border_style='thin'),
                             top=Side(border_style='thin'),
                             bottom=Side(border_style='thin'))

        font_normal = Font(name='Arial', size=8)

        alignment_left = Alignment(horizontal="left",
                                   vertical="top",
                                   wrap_text=True)

        for testcase in ast.literal_eval(self.jsondata['TestCaseList']):
            self._writeTestCaseTitles(line=self.line, testcase=testcase['name'])
            self.line += 2

            self._writeColumnTitles(line=self.line)
            self.line += 1

            for i in range(len(self.jsondata['TestCases'][testcase['name']])):

                style_range(self.xlws[chr(ord(cell_column) + 1) + str(self.line):
                                      chr(ord(cell_column) + 4) + str(self.line)],
                            border=border_thin,
                            font=font_normal,
                            alignment=alignment_left,
                            merge=False)

                self.xlws[chr(ord(cell_column) + 1) + str(self.line)] = \
                    self.jsondata['TestCases'][testcase['name']][i]["Step"]
                self.xlws[chr(ord(cell_column) + 2) + str(self.line)] = \
                    self.jsondata['TestCases'][testcase['name']][i]["Title"]
                self.xlws[chr(ord(cell_column) + 3) + str(self.line)] = \
                    self.jsondata['TestCases'][testcase['name']][i]["Description"]
                self.xlws[chr(ord(cell_column) + 4) + str(self.line)] = \
                    self.jsondata['TestCases'][testcase['name']][i]["ExpectedResults"]
                self.line += 1

            # TODO: add pagebreak between testtypes
            # self.xlws.page_breaks.append(Break(id=5))

            self.line += 1


if __name__ == "__main__":
    print 'Number of arguments:', len(sys.argv), 'arguments.'
    if len(sys.argv) > 3:
        raw_input("ERROR: Only one argument accepted.")
        exit(0)

    if len(sys.argv) == 1:
        jsonfile = raw_input("Enter jsonfile:")
        if jsonfile == "":
            jsonfile = "C:\\temp\\33.92.9999_P52.json"
        outputfolder = raw_input("Enter outputfolder:")
        if outputfolder == "":
            outputfolder = "C:\\temp"
    else:
        jsonfile = sys.argv[1]
        outputfolder = sys.argv[2]
        print 'jsonfile:', sys.argv[1]
        print 'outputfolder:', sys.argv[2]

    # jsonfile = "..\\TestFramework\\33.98.0156\\data\\33.92.9999_P52.json"
    # C:\\temp\\33.92.9999_P52.json
    TestSpecsXlsx(jsonfile=jsonfile, outputfolder=outputfolder).write()

    print("THE END")
