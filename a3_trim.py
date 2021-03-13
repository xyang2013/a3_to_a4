#!/usr/bin/python3
#

"""a3_trim

I am using HP OfficeJet Pro 7740. The size of scanned A3 documents are not
correct, 297.01 x 432.65mm instead of 297 x 420mm. This script trims off left
and bottom edges to make sure the pdf size is exactly A3 large.

Usage:
python ./a3_trim.py -i input.pdf -o output.pdf

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



import sys, getopt
from PyPDF2 import PdfFileWriter, PdfFileReader

# A3 dimensions
A3_WIDTH = 420 # in mm
A3_HEIGHT = 297 # in mm

# scanner scan area
SCAN_AREA_HEIGHT = 297.01 # in mm
SCAN_AREA_WIDTH = 432.65 # in mm

left_trim = SCAN_AREA_WIDTH - A3_WIDTH
bottom_trim = SCAN_AREA_HEIGHT - A3_HEIGHT

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

        # only process the first page
        first_page_number = 0

        first_page = input_.getPage(first_page_number)

        page_width = first_page.mediaBox.getUpperRight_x()
        page_height = first_page.mediaBox.getUpperRight_y()

        # use Py2PDF representation
        left_trim_ = left_trim/SCAN_AREA_WIDTH*float(page_width)
        bottom_trim_ = bottom_trim/SCAN_AREA_HEIGHT*float(page_height)

        first_page.cropBox.lowerLeft = (left_trim_, bottom_trim_)
        first_page.cropBox.upperRight = (page_width, page_height)

        output.addPage(first_page)
        
        with open(outputfile, "wb") as out_f:
            output.write(out_f)

if __name__ == "__main__":
    main(sys.argv[1:])