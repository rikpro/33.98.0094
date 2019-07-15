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
        self.jsonfile = jsonfile
        self.outputfolder = outputfolder
        
    def write(self):
        """
        Write complete xlsx file
        """
        base = os.path.basename(jsonfile)

        device = base.split("_P52")[0]
        print "device: ", device

        xlsxfile = os.path.join(os.path.normpath(self.outputfolder.strip('"')),
                                device + "_P52_1.01" + ".xlsx")

        if xlsxfile is not None:
            with open(xlsxfile, 'w') as outfile:
                outfile.write(str(datetime.now()))
            print "Saved XLSX file: %s" % xlsxfile
        

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
