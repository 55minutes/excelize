#!/usr/bin/env python
# -*- coding: utf-8 -*-
from ConfigParser import SafeConfigParser
from setuptools import setup, find_packages
import codecs
import os
import sys

version_config = SafeConfigParser()
version_config.readfp(open(
    os.path.join(os.path.dirname(__file__), 'VERSION.cfg')))
VERSION = version_config.get('version', 'working')

DISTRIBUTION_NAME = 'excelize'

SHORT_DESCRIPTION = 'Convenience wrapper for openpyxl'
if os.path.exists("README.md"):
    LONG_DESCRIPTION = codecs.open("README.md", "r", "utf-8").read()
else:
    LONG_DESCRIPTION = SHORT_DESCRIPTION

PROJECT_URL = 'https://github.com/55minutes/excelize'
DOWNLOAD_URL = '{0}/archive/master.tar.gz'.format(PROJECT_URL)

REQUIRES = ['clt-utils', 'distribute', 'openpyxl', 'unicodecsv']
if sys.version_info < (2, 7):
    REQUIRES.append('argparse')

# Setup the project directory
setup(
    name=DISTRIBUTION_NAME,
    version=VERSION,
    author='55 Minutes',
    author_email='info@55minutes.com',
    maintainer='55 Minutes',
    maintainer_email='info@55minutes.com',
    url=PROJECT_URL,
    description=SHORT_DESCRIPTION,
    long_description=LONG_DESCRIPTION,
    download_url=DOWNLOAD_URL,
    platforms=['any'],
    # See http://pypi.python.org/pypi?%3Aaction=list_classifiers
    classifiers=[
        'Development Status :: 5 - Production/Stable'
        'Environment :: Console',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: Apache Software License',
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Operating System :: POSIX',
        'Programming Language :: Python',
        'Topic :: Utilties',
    ],
    license='All Rights Reserved (c) 2012 55 Minutes',

    entry_points={
        'console_scripts': [
            'excelize = excelize.bin.excelize:main',
        ]
    },

    packages=find_packages(),
    zip_safe=False,
    include_package_data=True,

    dependency_links=[
        ('https://github.com/55minutes/clt-utils/archive/v1.0.1.tar.gz'
         '#egg=clt-utils-1.0.1'),
    ],
    install_requires=REQUIRES,
)
