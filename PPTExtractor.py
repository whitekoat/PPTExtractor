# -*- coding: utf-8 -*-
"""
Extract images from PowerPoint files


Extract images from PowerPoint files (ppt, pps, pptx) without use win32 API

required:
    OleFileIO_PL: Copyright (c) 2005-2010 by Philippe Lagadec

Usage

By default images are saved in current directory:

    ppt = PPTExtractor(file)

    # found images
    len(ppt)

    # extract image
    ppt.extract(images[0])

    # extract all images
    ppt.extractall()

"""
# Copyright (c) 2010 Jhonathan Salguero Villa (http://github.com/sney2002)
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

import os
import struct
import zipfile
from io import BytesIO

import OleFileIO_PL as OleFile

DEBUG = False
CWD = '.'
CHUNK = 1024 * 64

# MS-ODRAW spec
formats = {
    # 2.2.24
    (0xF01A, 0x3D40): (50, ".emf"),
    (0xF01A, 0x3D50): (66, ".emf"),
    # 2.2.25
    (0xF01B, 0x2160): (50, ".wmf"),
    (0xF01B, 0x2170): (66, ".wmf"),
    # 2.2.26
    (0xF01C, 0x5420): (50, ".pict"),
    (0xF01C, 0x5430): (50, ".pict"),
    # 2.2.27
    (0xF01D, 0x46A0): (17, ".jpeg"),
    (0xF01D, 0x6E20): (17, ".jpeg"),
    (0xF01D, 0x46B0): (33, ".jpeg"),
    (0xF01D, 0x6E30): (33, ".jpeg"),
    # 2.2.28
    (0xF01E, 0x6E00): (17, ".png"),
    (0xF01E, 0x6E10): (33, ".png"),
    # 2.2.29
    (0xF01F, 0x7A80): (17, ".dib"),
    (0xF01F, 0x7A90): (33, ".dib"),
    # 2.2.30
    (0xF029, 0x6E40): (17, ".tiff"),
    (0xF029, 0x6E50): (33, ".tiff")
}


class InvalidFormat(Exception):
    pass


class PowerPointFormat(object):
    def __init__(self, file):
        """
        filename:   archivo a abrir
        """
        self._files = []

        self._process(file)

    def extract(self, index):
        """
        Extrae imagen en directorio especificado.
        """
        return self._extract(index)

    def extractall(self):
        """
        Extrae todas las imágenes en directorio especificado.
        """
        for index in range(len(self._files)):
            yield self.extract(index)

    def __len__(self):
        return len(self._files)

    def __str__(self):
        return "<PowerPoint file with %s images>" % len(self)

    __repr__ = __str__


# TODO: Extraer otros tipos de archivo (wav, avi...)
class PPT(PowerPointFormat):
    """
    Extrae imágenes de archivos PowerPoint binarios (ppt, pps).
    """
    headerlen = struct.calcsize('<HHL')

    @classmethod
    def is_valid_format(cls, file):
        return OleFile.isOleFile(file)

    def _process(self, file):
        """
        Busca imágenes dentro de stream y guarda referencia a su ubicación.
        """
        olefile = OleFile.OleFileIO(file)

        # Al igual que en pptx esto no es un error
        if not olefile.exists("Pictures"):
            return
            # raise IOError("Pictures stream not found")

        self.__stream = olefile.openstream("Pictures")

        stream = self.__stream
        offset = 0
        # cantidad de imágenes encontradas
        n = 1

        while True:
            header = stream.read(self.headerlen)
            offset += self.headerlen

            if not header:
                break

            # cabecera
            recInstance, recType, recLen = struct.unpack_from("<HHL", header)

            # mover a siguiente cabecera
            stream.seek(recLen, 1)

            if DEBUG:
                print("%X %X %sb" % (recType, recInstance, recLen))

            extrabytes, ext = formats.get((recType, recInstance))

            # Eliminar bytes extra
            recLen -= extrabytes
            offset += extrabytes

            self._files.append((offset, recLen))
            offset += recLen

            n += 1

    def _extract(self, index):
        """
        Extrae imagen en el directorio actual (path).
        """
        if index >= len(self._files):
            raise IOError("No such file")

        offset, size = self._files[index]

        total = 0

        self.__stream.seek(offset, 0)

        output = BytesIO()
        while (total + CHUNK) < size:
            data = self.__stream.read(CHUNK)

            if not data:
                break

            output.write(data)
            total += len(data)

        if total < size:
            data = self.__stream.read(size - total)
            output.write(data)
        return output


class PPTX(PowerPointFormat):
    """
    Extrae imágenes de archivos PowerPoint +2007
    """

    @classmethod
    def is_valid_format(cls, file):
        return zipfile.is_zipfile(file)

    def _process(self, file):
        """
        Busca imágenes dentro de archivo zip y guarda referencia a su ubicación
        """
        self.__zipfile = zipfile.ZipFile(file)

        n = 1

        for file in self.__zipfile.namelist():
            path, name = os.path.split(file)
            name, ext = os.path.splitext(name)

            # los archivos multimedia se guardan en ppt/media
            if path == "ppt/media":
                # guardar path de archivo dentro del zip
                self._files[n] = file

                n += 1

    def _extract(self, index):
        """
        Extrae imagen en el directorio actual (path).
        """
        if index >= len(self._files):
            raise IOError("No such file")

        total = 0

        # extraer archivo
        file = self.__zipfile.open(self._files[index])

        output = BytesIO()
        while True:
            data = file.read(CHUNK)

            if not data:
                break

            output.write(data)
            total += len(data)
        return output


def PPTExtractor(file):
    # Identificar tipo de archivo (pps, ppt, pptx) e instanciar clase adecuada
    for cls in PowerPointFormat.__subclasses__():
        if cls.is_valid_format(file):
            return cls(file)
    raise InvalidFormat("{0} is not a PowerPoint file".format(file))
