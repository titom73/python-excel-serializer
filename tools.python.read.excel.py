#/usr/bin/env python

""" Extract data from EXCEL to fill-in JINJA2 template

This script is a tool to convert data from Excel file to YAML to fill-in a Jinja2 template.

YAML structure is the following:
SHEET_NAME:
  - COLUMN#1: value ROW#1 / COLUMN#1
    COLUMN#2: value ROW#1 / COLUMN#2
    ...
  - COLUMN#1: value ROW#2 / COLUMN#1
    COLUMN#2: value ROW#2 / COLUMN#2
    ...

Usage:

python python-tools/tools.python.read.excel.py -e topology.xlsx -s Topology -t topology.j2
 * Reading excel file: topology.xlsx
 * Template rendering
    > Use template: topology.j2
    > Output file: output.txt

Keyword arguments:

excel -- path to excel file to read data
sheet -- Sheet name within Excel file to extract data
template -- Jinja2 template to use for rendering
output -- File to create for the rendering

Return Value:

Text file
"""

import argparse
import os.path
import yaml
from jinja2 import Template
import urllib2

from jtools import *
from jtools.excel import *


def download_file( remote_file=None, verbose=False):
	url = remote_file
	file_name = url.split('/')[-1]
	u = urllib2.urlopen(url)
	f = open(file_name, 'wb')
	meta = u.info()
	file_size = int(meta.getheaders("Content-Length")[0])
	if verbose is True:
		print " * Downloading: %s Bytes: %s" % (file_name, file_size)
	file_size_dl = 0
	block_sz = 8192
	while True:
	    buffer = u.read(block_sz)
	    if not buffer:
	        break
	    file_size_dl += len(buffer)
	    f.write(buffer)
	    if verbose is True:
		    status = r"%10d  [%3.2f%%]" % (file_size_dl, file_size_dl * 100. / file_size)
		    status = status + chr(8)*(len(status)+1)
		    print "    > "+status,
	f.close()
	return file_name


if __name__ == "__main__":

	### CLI Option parser:
	parser = argparse.ArgumentParser(description="Excel to text file ")
	parser.add_argument('-e', '--excel', help='Input Excel file',default=None)
	parser.add_argument('-u', '--url', help='URL for remote Excel file',default=None)
	parser.add_argument('-s', '--sheet', help='Excel sheet tab',default="Sheet1")
	parser.add_argument('-m', '--mode', help='Serializer mode: table / list',default="table")
	parser.add_argument('-n', '--nb_columns', help='Number of columns part of the table',default=6)
	parser.add_argument('-t', '--template', help='Template file to render',default=None)
	parser.add_argument('-o', '--output', help='Output file ',default="output.txt")
	parser.add_argument('-v', '--verbose',action='store_true', help='Increase verbositoy for debug purpose',default=False)
	options = parser.parse_args()

	# Start Engine checklist
	if options.excel is None and options.url is None:
		sys.exit("Excel file is missing please provide it with option -e or -u")
	elif options.excel is not None and options.url is not None:
		sys.exit("These options are mutually exclusive, please use -e OR -u")
	elif options.sheet is None:
		sys.exit("Sheet tab is not define, please provide it with option -s")
	# elif options.template is None:
	# 	sys.exit("Jinja 2 Template file is missing please provide it with option -t")

	# Download Excel file from Webserver:
	if options.url is not None:
		print " * Downloading Excel file from "+ options.url
		options.excel = download_file(remote_file= options.url, verbose= options.verbose)

	# Serialize EXCEL file
	print " * Reading excel file: "+options.excel
	table = ExcelSerializer(excel_path=options.excel)
	if options.mode == "table":
		table.serialize_table( sheet=options.sheet, nb_columns=options.nb_columns)
	elif options.mode == "list":
		table.serialize_list( sheet=options.sheet)
	else:
		sys.exit("Error: unsuported serializer mode")

	if options.verbose is True:
		print " ** Debug Output **"
		if table.get_yaml(sheet=options.sheet) is None:
			print "    -> Sheet has not been found in the structure"
		else:
			print yaml.safe_dump(table.get_yaml(sheet=options.sheet), encoding='utf-8', allow_unicode=True)
		print " ** Debug Output **"

	# Jinja2 template file.
	# Open JINJA2 file as a template
	if options.template is not None and options.output is not None:
		print " * Template rendering"
		print "    > Use template: "+ options.template
		print "    > Output file: "+ options.output

		with open( options.template ) as t_fh:
			t_format = t_fh.read()
		template = Template( t_format )
		# Create output file 
		confFile = open( options.output,'w')
		confFile.write( template.render( table.get_yaml() ) )
		confFile.close()
	else:
		if table.get_yaml(sheet=options.sheet) is None:
			print "    -> Sheet has not been found in the structure"
		else:
			print yaml.safe_dump(table.get_yaml_all(), encoding='utf-8', allow_unicode=True)
