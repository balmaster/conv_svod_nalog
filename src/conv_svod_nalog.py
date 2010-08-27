#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os

import sys

from optparse import OptionParser

from Ft.Xml.Xslt import Processor
# We use the InputSource architecture
from Ft.Xml import InputSource
from Ft.Lib.Uri import OsPathToUri  # path to URI conversions
from Ft.Xml import Parse
import Ft.Xml.Domlette
import string
import re

def convert(options):


    m = re.search('(\d\w+)(\d{4})', options.inputFile)
    form = m.group(1)
    year = m.group(2)


    processor = Processor.Processor()

    srcAsUri = OsPathToUri(options.inputFile)
    source = InputSource.DefaultFactory.fromUri(srcAsUri)

    ssAsUri = OsPathToUri("prepare_data.xsl")
    transform = InputSource.DefaultFactory.fromUri(ssAsUri)

    processor.appendStylesheet(transform)

    xml = processor.run(source)

#    out = open('out','w')
#    out.write(xml)
#    out.close()

    doc  = Parse(xml)

    reportList = doc.xpath('reports/report')
    i = 0
    for report in reportList:
        i=i+1
        print 'report:' , i
        name = report.xpath('string(@name)')
        count = report.xpath('string(@count)')

        

        dataFile = open(''.join([form,year,'_',str(i),'.csv']),'w')

        for p in report.xpath(''.join(['p[count(pv)=',count,']'])):
      
            line = ''.join([form,';',year,';',name,';'])
            for pv in p.xpath('pv'):
                value = pv.xpath('string(@value)')
                value = string.replace(value,'\n','')
                value = string.strip(value,' ')
                line = ''.join([line,value,';'])

            dataFile.write(line.encode('cp1251'))
            dataFile.write('\n')


        dataFile.close()
    return


parser = OptionParser()
parser.add_option("-i", "--input-file", action="store", dest="inputFile", help="input file name", default=None)

(options, args) = parser.parse_args()

if not options.inputFile:
    print "input file is required!"
    parser.print_help()
    sys.exit(1)

convert(options)

sys.exit(0)
