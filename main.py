#!/usr/bin/python
#

"""A3 to A4

This script converts an A3 pdf (in landscape) to two A4 pdf pages (in portrait) without information loss in the center region. The A3 size print measures 29.7 x 42.0cm. The A4 size print measures 21.0 x 29.7cm. The program sacrifices the left and right edges of the A3 document to make sure there is no information loss in the middle region due to a printer' physical print margin.  

For example, the minimum printing margin of my HP LaserJet Professional P1606dn is 4mm (based on actual measurements). I use 4mm for the overlapping parameter. An A3 pdf document is split into two A4 pdf documents. The first A4 page maps the area between 4mm and 214mm (210+4) of the A3 document. The second A4 page maps the area between 206mm (210-4) and 416mm (420-4) of the A3 document. Hence, when we print the two A4 documents given the physical printing margin, there should not be any information loss in the middle if we join the two pages together. You should adjust the overlapping parameter for your printer.

To make it easy to use, you can follow the following steps:
1. Download and install Anaconda
2. Install PyPDF2 package
3. Rename the A3 pdf that you want to split to in.pdf
4. Place this script in the same folder as the pdf
5. Run this Python script

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


from PyPDF2 import PdfFileWriter, PdfFileReader
import copy

# dimensions in landscape
A3_WIDTH = 42. # in cm
A3_HEIGHT = 29.7 # in cm
A4_WIDTH = 29.7 # in cm
A4_HEIGHT = 21. # in cm

overlapping = 0.4 # in cm


with open("in.pdf", "rb") as in_f:
    input1 = PdfFileReader(in_f)
    output = PdfFileWriter()

    numPages = input1.getNumPages()
    print("document has %s pages." % numPages)

    for i in range(numPages):

        # prepare page 1

        page_first_half = copy.copy(input1.getPage(i))
        page_width = page_first_half.mediaBox.getUpperRight_x()
        page_height = page_first_half.mediaBox.getUpperRight_y()

        a4_1_starting = overlapping/A3_WIDTH*float(page_width)
        a4_1_ending = (overlapping+A4_HEIGHT)/A3_WIDTH*float(page_width)

        page_first_half.cropBox.lowerLeft = (a4_1_starting, 0)
        page_first_half.cropBox.upperRight = (a4_1_ending, page_height)
        output.addPage(page_first_half)

        # prepare page 2

        page_second_half = copy.copy(input1.getPage(i))

        a4_2_starting = (A4_HEIGHT-overlapping)/A3_WIDTH*float(page_width)
        a4_2_ending = (A3_WIDTH-overlapping)/A3_WIDTH*float(page_width)

        page_second_half.cropBox.lowerLeft = (a4_2_starting, 0)
        page_second_half.cropBox.upperRight = (a4_2_ending, page_height)
        output.addPage(page_second_half)

    with open("out.pdf", "wb") as out_f:
        output.write(out_f)