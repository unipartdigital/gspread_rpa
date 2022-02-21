from setuptools import setup
import os, sys

sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src'))
from gspread_rpa import __version__

pkg_data = [os.path.join('demo', '*demo*.py'),
            os.path.join('demo', 'Makefile'),
            os.path.join('demo', 'make.bat')]

setup(
    name='gspread-rpa',
    version=__version__,
    description='a gspread (Python API for Google Sheets) hight level wrapper',
    author='Unipart Digital',
    author_email='rpa@unipart.io',
    maintainer='Ali Bendriss',
    maintainer_email='ali.bendriss@gmail.com',
    url='https://github.com/unipartdigital/gspread_rpa',
    packages=['gspread_rpa'],
    package_dir={'gspread_rpa': os.path.join('src', 'gspread_rpa')},
    package_data={'': pkg_data},
    requires=['gspread'],
    license='GPLv3',
    classifiers=[
        'Environment :: Console',
        'Environment :: Web Environment',
        'Intended Audience :: End Users/Desktop',
        'Intended Audience :: Developers',
        'Intended Audience :: System Administrators',
        'Operating System :: MacOS :: MacOS X',
        'Operating System :: Microsoft :: Windows',
        'Operating System :: POSIX',
        'Programming Language :: Python',
        'Topic :: Communications :: Email',
        'Topic :: Office/Business',
        'Topic :: Software Development :: Bug Tracking',
    ],
)
