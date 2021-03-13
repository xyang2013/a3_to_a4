#!/usr/bin/python
#

"""a3_to_a4

This script converts an A3 pdf (in landscape) to two A4 pdf pages (in portrait)
without any information loss in the center region. The A3 size print measures 
29.7 x 42.0cm. The A4 size print measures 21.0 x 29.7cm. The program sacrifices
the left and right edges of the A3 pdf to make sure there is no information loss
in the central region which is usually the case because of the physical printing
margins of a printer. For example, the minimum printing margin of my HP LaserJet
Professional P1606dn is 4mm (measured with a ruler). To compensate for the 
effect of the margin in the center region, when an A3 pdf document is split into
two A4 pdf documents, the first A4 page maps the area between 4mm and 214mm 
(210+4) of the A3 document, and the second A4 page maps the area between 206mm
(210-4) and 416mm (420-4) of the A3 document. You should adjust the 
'overlapping' parameter for your printer.

Requirements: 
Python3, PyPDF2

Usage:
python ./a3_to_a4.py -i ./input_a3.pdf -o ./input_a3_to_a4.pdf

Copyright (c) 2021 Xiao Yang

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

"""

__author__ = "Xiao Yang"
__copyright__ = "Copyright (C) 2021 Xiao Yang"
__license__ = "MIT License"
__version__ = "1.0"

import sys, getopt, copy
from PyPDF2 import PdfFileWriter, PdfFileReader

# A3 dimensions
A3_WIDTH = 420 # in mm
A3_HEIGHT = 297 # in mm

# A4 dimensions
A4_WIDTH = 297 # in mm
A4_HEIGHT = 210 # in mm

# calibrate the following parameter according to your printer
overlapping = 4 # in mm

def main(argv):

    inputfile = ''
    outputfile = ''
    try:
        opts, _ = getopt.getopt(argv,"hi:o:",["ifile=","ofile="])
    except getopt.GetoptError:
        print ('a3_trim.py -i <inputfile> -o <outputfile>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print ('test.py -i <inputfile> -o <outputfile>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
        elif opt in ("-o", "--ofile"):
            outputfile = arg
    print ('Input file is "', inputfile)
    print ('Output file is "', outputfile)

    with open(inputfile, "rb") as in_f:
        input_ = PdfFileReader(in_f)
        output = PdfFileWriter()

        numPages = input_.getNumPages()
        print("document has %s pages." % numPages)

        for i in range(numPages):

            # prepare page 1

            page_first_half = copy.copy(input_.getPage(i))
            page_width = page_first_half.mediaBox.getUpperRight_x()
            page_height = page_first_half.mediaBox.getUpperRight_y()

            a4_1_starting = overlapping/A3_WIDTH*float(page_width)
            a4_1_ending = (overlapping+A4_HEIGHT)/A3_WIDTH*float(page_width)

            page_first_half.cropBox.lowerLeft = (a4_1_starting, 0)
            page_first_half.cropBox.upperRight = (a4_1_ending, page_height)
            output.addPage(page_first_half)

            # prepare page 2

            page_second_half = copy.copy(input_.getPage(i))

            a4_2_starting = (A4_HEIGHT-overlapping)/A3_WIDTH*float(page_width)
            a4_2_ending = (A3_WIDTH-overlapping)/A3_WIDTH*float(page_width)

            page_second_half.cropBox.lowerLeft = (a4_2_starting, 0)
            page_second_half.cropBox.upperRight = (a4_2_ending, page_height)
            output.addPage(page_second_half)

        with open(outputfile, "wb") as out_f:
            output.write(out_f)

if __name__ == "__main__":
    main(sys.argv[1:])
