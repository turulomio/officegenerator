from setuptools import setup, Command
import site
import os
import platform


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
  * Cambiar la versión y la fecha en version.py
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
        os.system("rm -Rf build")
        os.system("rm -Rf doc/html")
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

with open('README.rst', encoding='utf-8') as f:
    long_description = f.read()

## Version of officegenerator captured from commons to avoid problems with package dependencies
__version__= None
with open('officegenerator/commons.py', encoding='utf-8') as f:
    for line in f.readlines():
        if line.find("__version__ =")!=-1:
            __version__=line.split("'")[1]


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
     install_requires=['odfpy','openpyxl'],
     entry_points = {'console_scripts': ['officegenerator_demo=officegenerator.demo:main',
                                         'officegenerator_odf2xml=officegenerator.odf2xml:main',
                                        ],
                    },
     cmdclass={'doxygen': Doxygen,
               'uninstall':Uninstall, 
               'doc': Doc,
               'procedure': Procedure,
              },
     zip_safe=False,
     test_suite = 'officegenerator.tests',
     include_package_data=True
)
