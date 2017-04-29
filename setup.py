import os

from setuptools import setup, find_packages


here = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(here, 'README.md')) as f:
    README = f.read()

requires = [
    'olefile'
    ]

setup(name='PPTExtractor',
      py_modules=['PPTExtractor'],
      install_requires=requires,
      )
