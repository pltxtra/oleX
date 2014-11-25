#!/usr/bin/python

# (C) 2014 by Anton Persson

# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
#(at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

import pyolecf
import sys
import tempfile
import zipfile
import os
import shutil

def extractOleItem(item, path):
    fo = open(path, "wb")
    print "creating file: " + path
    size = item.get_size()
    content = item.read(size)
    fo.write(content)
    fo.close()
    
def extractOleFile(file_name,destination):
    olecf_file = pyolecf.file()
    olecf_file.open(file_name)
    root = olecf_file.get_root_item()
    for x in range(0, root.get_number_of_sub_items()):
        sub_item = root.get_sub_item(x)
        if sub_item.get_name() == "CONTENTS":
            extractOleItem(item=sub_item, path=destination)
    olecf_file.close()

TDIR=tempfile.mkdtemp(suffix="_oleX")

ZIPFILE = zipfile.ZipFile(sys.argv[1], "r")
ZIPFILE.extractall(path=TDIR)

EPTH=TDIR + "/word/embeddings"
files = os.listdir(EPTH)

for file in files:
    extractOleFile(file_name=EPTH + "/" + file, destination = file + ".extracted")
    
shutil.rmtree(TDIR)
