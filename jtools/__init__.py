import pprint
import os
import jinja2
import colorful
from glob import glob

def jprint( data ):
	""" Display Python data struccture using pprint
	Useful shortcut to display any content without needs to instantiate PPRINT

	Keyword arguments:
	data -- Variable to display using pprint
	"""
	pp = pprint.PrettyPrinter(indent=4)
	pp.pprint(data)

def get_list_files( path=None, extension="j2"):
	""" Get a recursive list of file with a specific extension.

	Keyword arguements:
	path -- Root directory to look for files
	extension -- File extension to look for. Default is j2

	Return Value:
	result -- A Python list of file with relative path to access to it
	"""
	result = [y for x in os.walk(path) for y in glob(os.path.join(x[0], '*.'+extension))]
	return result

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