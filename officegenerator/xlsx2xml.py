## @namespace officegenerator.odf2xml
## @brief Function to convert a ODF file to xml output
## 
## Executes the following code in bash
## @code
## odf2xml -o "$1.no.xml" "$1"
## cat "$1.no.xml" | xmllint --format -
## @endcode

import argparse
import gettext
import pkg_resources
import platform
import subprocess
import sys

from officegenerator.commons import argparse_epilog,  __version__
from os import system
from zipfile import ZipFile

try:
    t=gettext.translation('officegenerator',pkg_resources.resource_filename("officegenerator","locale"))
    _=t.gettext
except:
    _=str

## Generates an xml file with the filename+".xml" name
def generate_xml(filename):
    zipfilename=filename+".zip"
    system("cp '{}' '{}'".format(filename, zipfilename))
    zf = ZipFile(zipfilename)
    for fi in zf.namelist():
        print("### {} ###".format(fi))
        xml=open("tmp.xml", "w")
        print(fi)
        xml.write(zf.read(fi).decode('UTF-8'))
        xml.close()
        q=subprocess.run(['xmllint', '--format', 'tmp.xml'])
        if q.stdout!=None:
            print(q.stdout.decode('UTF-8'))
        subprocess.run(['rm', 'tmp.xml'])
        print("")
    subprocess.run(['rm', zipfilename])

def main():
    parser=argparse.ArgumentParser(prog='officegenerator', description=_('Convert a XLSX file into an indented xml'), epilog=argparse_epilog(), formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--file', action='store', help=_('XLSX file to convert'), required=True)
    args=parser.parse_args()

    if platform.system()!="Linux":
        print(_("officegenerator_xlsx2xml only works on Linux"))
        sys.exit(1)

    generate_xml(args.file)

if __name__ == "__main__":
    main()
