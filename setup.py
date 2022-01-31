from setuptools import setup
from os import path

setup(
    name='gspread-rpa',
    version='1.0.0',
    description='Python gspread google spreadsheet API wrapper',
    author='Unipart Digital',
    author_email='rpa@unipart.io',
    maintainer='Ali Bendriss',
    maintainer_email='ali.bendriss@gmail.com',
    url='https://github.com/unipart',
    packages=['gspread_rpa'],
    package_dir={'gspread_rpa': path.join('src', 'gspread_rpa')},
    package_data={'': [path.join('demo', '*demo*.py'), path.join('demo', 'Makefile')]},
    requires=['gspread'],
    license='GPLv3',
    classifiers=[
        'Development Status :: 42 - Beta',
        'Environment :: Console',
        'Environment :: Web Environment',
        'Intended Audience :: End Users/Desktop',
        'Intended Audience :: Developers',
        'Intended Audience :: System Administrators',
        'License :: OSI Approved :: GNU GENERAL PUBLIC LICENSE Version 3',
        'Operating System :: MacOS :: MacOS X',
        'Operating System :: Microsoft :: Windows',
        'Operating System :: POSIX',
        'Programming Language :: Python',
        'Topic :: Communications :: Email',
        'Topic :: Office/Business',
        'Topic :: Software Development :: Bug Tracking',
    ],
)
