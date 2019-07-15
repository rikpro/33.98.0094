#!/usr/bin/python2.7
"""
Generate Dummy JSON files
"""
__author__ = "Rik Provoost"
__company__ = "Televic Rail"

import sys
import os
import pydoc
import ast
import logging
import json
import datetime


DEBUG = False

class TestSpecification(object):
    """
    Create Test Specifications Class
    """
    try:
        LOCAL_LOCATION = os.path.dirname(os.path.realpath(__file__)) + "\\"
    except:
        LOCAL_LOCATION = os.path.dirname(os.path.realpath(sys.argv[0])) + "\\"

    if DEBUG is True:
        DATA_FOLDER = "C:\\temp\\"
    else:
        DATA_FOLDER = LOCAL_LOCATION + "data\\"
    print "JSON DATA FILES FOLDER: ", DATA_FOLDER

    def __init__(self, partnumber=None):

        if partnumber is None:
            self.partnumber = '%'
        else:
            self.partnumber = partnumber
        
        self.testspecrecords = None
        self.jsondata = None

    def run(self):
        """
        Generate testspecifications with export to JSON file.
        """

        jsonfile = self._exportJson(device=self.partnumber)

    def _exportJson(self, device=None, records=None):
        """
        Export JSON data to JSON file.

        :param device: dict representing the device {dict['Name'], dict['Description']}
        :type device:  dict
        :param records: testspecifications
        :type records: list of dicts
        :return: JSON filename
        :rtype: str
        """
        # Save JSON to file
        jsonfile = os.path.join(TestSpecification.DATA_FOLDER,
                                device + "_P52" + ".json")

        if jsonfile is not None:
            with open(jsonfile, 'w') as outfile:
                outfile.write(str(datetime.datetime.now()))
            print "Saved JSON file: %s" % jsonfile
        return jsonfile


if __name__ == '__main__':
    print 'Number of arguments:', len(sys.argv), 'arguments.'
    if len(sys.argv) > 2:
        raw_input("ERROR: Only one argument accepted.")
        exit(0)

    if len(sys.argv) == 1:
        # if partnumber is not entered, all partnumbers will be used
        partnumber = raw_input("Enter partnumber:")
    else:
        partnumber = sys.argv[1]
        print 'Argument :', sys.argv[1]

    # get data and output
    TestSpecification(partnumber=partnumber).run()

    # #DEBUG
    # partnumber = "33.92.9999"
    # test = TestSpecification(partnumber=partnumber)
    # testtypes = test.getDatabaseTestTypes()
    # print "testtypes: ", testtypes
