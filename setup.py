from setuptools import setup, Command
import site
import os
import platform



class Reusing(Command):
    description = "Download modules from https://github.com/turulomio/reusingcode/"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        from sys import path
        path.append("officegenerator")
        from github import download_from_github
        download_from_github('turulomio','reusingcode','python/github.py', 'officegenerator')
        download_from_github('turulomio','reusingcode','python/casts.py', 'officegenerator')
        download_from_github('turulomio','reusingcode','python/datetime_functions.py', 'officegenerator')
        download_from_github('turulomio','reusingcode','python/decorators.py', 'officegenerator')
        download_from_github('turulomio','reusingcode','python/libmanagers.py', 'officegenerator')
        download_from_github('turulomio','reusingcode','python/objects/percentage.py', 'officegenerator/objects/')
        download_from_github('turulomio','reusingcode','python/objects/currency.py', 'officegenerator/objects/')

## Class to define doc command
class Doc(Command):
    description = "Update translations"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        #es
        os.system("xgettext -L Python --no-wrap --no-location --from-code='UTF-8' -o locale/officegenerator.pot *.py officegenerator/*.py")
        os.system("msgmerge -N --no-wrap -U locale/es.po locale/officegenerator.pot")
        os.system("msgfmt -cv -o officegenerator/locale/es/LC_MESSAGES/officegenerator.mo locale/es.po")

class Procedure(Command):
    description = "Show release procedure"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        print("""Nueva versión:
  * Cambiar la versión y la fecha en commons.py
  * Modificar el Changelog en README
  * python setup.py doc
  * linguist
  * python setup.py doc
  * python setup.py install
  * python setup.py doxygen
  * git commit -a -m 'officegenerator-{}'
  * git push
  * Hacer un nuevo tag en GitHub
  * python setup.py sdist upload -r pypi
  * python setup.py uninstall
  * Crea un nuevo ebuild de Gentoo con la nueva versión
  * Subelo al repositorio del portage
""".format(__version__))

## Class to define doxygen command
class Doxygen(Command):
    description = "Create/update doxygen documentation in doc/html"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        print("Creating Doxygen Documentation")
        os.system("""sed -i -e "41d" doc/Doxyfile""")#Delete line 41
        os.system("""sed -i -e "41iPROJECT_NUMBER         = {}" doc/Doxyfile""".format(__version__))#Insert line 41
        os.system("rm -Rf build")
        os.chdir("doc")
        os.system("doxygen Doxyfile")
        os.system("rsync -avzP -e 'ssh -l turulomio' html/ frs.sourceforge.net:/home/users/t/tu/turulomio/userweb/htdocs/doxygen/officegenerator/ --delete-after")
        os.chdir("..")

## Class to define uninstall command
class Uninstall(Command):
    description = "Uninstall installed files with install"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        if platform.system()=="Linux":
            os.system("rm -Rf {}/officegenerator*".format(site.getsitepackages()[0]))
            os.system("rm /usr/bin/officegenerator*")
        else:
            os.system("pip uninstall officegenerator")

########################################################################

## Version of officegenerator captured from commons to avoid problems with package dependencies
__version__= None
with open('officegenerator/commons.py', encoding='utf-8') as f:
    for line in f.readlines():
        if line.find("__version__ =")!=-1:
            __version__=line.split("'")[1]


setup(name='officegenerator',
     version=__version__,
     description='Python module to read and write LibreOffice and MS Office files',
     long_description='Project web page is in https://github.com/turulomio/officegenerator',
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
     install_requires=['odfpy==1.3.6','openpyxl'],
     entry_points = {'console_scripts': ['officegenerator_demo=officegenerator.demo:main',
                                         'officegenerator_odf2xml=officegenerator.odf2xml:main',
                                         'officegenerator_xlsx2xml=officegenerator.xlsx2xml:main',
                                        ],
                    },
     cmdclass={'doxygen': Doxygen,
               'uninstall':Uninstall, 
               'doc': Doc,
               'procedure': Procedure,
               'reusing': Reusing,
              },
     zip_safe=False,
     test_suite = 'officegenerator.tests',
     include_package_data=True
)
