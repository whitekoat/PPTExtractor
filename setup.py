import os

from setuptools import setup


here = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(here, 'README.md')) as f:
    README = f.read()

requires = [
    'olefile == 0.46'
    ]

setup(name='PPTExtractor',
      version='0.0.4',
      py_modules=['PPTExtractor'],
      install_requires=requires,
      )
