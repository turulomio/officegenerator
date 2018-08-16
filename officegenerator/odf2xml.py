## @namespace officegenerator.odf2xml
## @brief Function to convert a ODF file to xml output
## 
## Executes the following code in bash
## @code
## odf2xml -o "$1.no.xml" "$1"
## cat "$1.no.xml" | xmllint --format -
## @endcode


import argparse
import datetime
import os
import subprocess

from .__init__ import __version__, __versiondate__

def main():
    parser=argparse.ArgumentParser(prog='officegenerator', description='Convert a ODF file into an indented xml', epilog="Developed by Mariano Muñoz 2018-{}".format(__versiondate__.year), formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--file', action='store', help='Odf file to convert')
    args=parser.parse_args()

    p=subprocess.run(["odf2xml", "-o", args.file + ".xml", args.file], stdout=subprocess.PIPE)
    q=subprocess.run(['xmllint', '--format', args.file +".xml"], stdout=subprocess.PIPE, input=p.stdout)
    r=subprocess.run(['rm',args.file+".xml"])
    print(q.stdout.decode('UTF-8'))

#odf2xml -o "$1.no.xml" "$1"
#cat "$1.no.xml" | xmllint --format -


if __name__ == "__main__":
    main()