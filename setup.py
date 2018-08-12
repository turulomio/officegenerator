from officegenerator import __version__
from setuptools import setup, Command

import datetime
import gettext
import os
import site

class Doxygen(Command):
    description = "Create/update doxygen documentation in doc/html"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        print("Creating Doxygen Documentation")
        os.chdir("doc")
        os.system("doxygen Doxyfile")
        os.system("rsync -avzP -e 'ssh -l turulomio' html/ frs.sourceforge.net:/home/users/t/tu/turulomio/userweb/htdocs/doxygen/officegenerator/ --delete-after")
        os.chdir("..")

class Uninstall(Command):
    description = "Uninstall installed files with install"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        os.system("rm -Rf {}/officegenerator*".format(site.getsitepackages()[0]))

########################################################################

with open('README.rst', encoding='utf-8') as f:
    long_description = f.read()

setup(name='officegenerator',
     version=__version__,
     description='Python module to read and write LibreOffice and MS Office files',
     long_description=long_description,
     long_description_content_type='text/markdown',
     classifiers=['Development Status :: 4 - Beta',
                  'Intended Audience :: Developers',
                  'Topic :: Software Development :: Build Tools',
                  'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
                  'Programming Language :: Python :: 3',
                 ], 
     keywords='office generator',
     url='https://officegenerator.sourceforge.io/',
     author='Turulomio',
     author_email='turulomio@yahoo.es',
     license='GPL-3',
     packages=['officegenerator'],
     cmdclass={
        'doxygen': Doxygen,
        'uninstall':Uninstall, 
             },
      zip_safe=False
     )
