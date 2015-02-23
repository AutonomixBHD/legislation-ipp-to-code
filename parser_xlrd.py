#! /usr/bin/env python
# -*- coding: utf-8 -*-


# OpenFisca -- A versatile microsimulation software
# By: OpenFisca Team <contact@openfisca.fr>
#
# Copyright (C) 2011, 2012, 2013, 2014, 2015 OpenFisca Team
# https://github.com/openfisca
#
# This file is part of OpenFisca.
#
# OpenFisca is free software; you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as
# published by the Free Software Foundation, either version 3 of the
# License, or (at your option) any later version.
#
# OpenFisca is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.


import xlrd


def handle_format_with_xlrd(path,bareme):
    book = xlrd.open_workbook(filename = path+ bareme +'.xls',formatting_info=True)
    sheet_names = book.sheet_names()
    print 'sheet_names',sheet_names
    sheet = book.sheet_by_name('BMAF')
#    for sheet in sheet_names:

    # travail sur une feuille d'abord
    #    for col in range(sheet.ncols):
    #            for row in range(sheet.nrows):

        # travail sur une cellule d'abord
    row = 40
    col = 1
    cell_type = sheet.cell_type(row,col)
    cell_value = sheet.cell_value(row,col)
    print 'type',cell_type, 'value',cell_value
    xf_index = sheet.cell_xf_index(row,col)
    xf = book.xf_list[xf_index]
    bgx = xf.background.pattern_colour_index
    print 'bgx',bgx
    format = book.format_map[xf.format_key] # gets a Format object
    format_str = format.format_str # this is the "number format string"

    print 'format ', format, format_str

    # Applcation de la fonction d'emmanuel
    format = format_xls_cell(book= book, type = cell_type, value = cell_value, xf_index  =xf_index)
    print 'format2',format
    dic_sheetname_by_format_matrix = {}
    return dic_sheetname_by_format_matrix

def format_xls_cell(book, type, value, xf_index): #type = cell_type ; value = cell_value
    """Convert an XLS cell (type & value) to an unicode string.

    Code taken from http://code.activestate.com/recipes/546518-simple-conversion-of-excel-files-into-csv-and-yaml/

    Type Codes:
    EMPTY   0
    TEXT    1 a Unicode string
    NUMBER  2 float
    DATE    3 float
    BOOLEAN 4 int; 1 means TRUE, 0 means FALSE
    ERROR   5
    """
    if type == 0:
        value = None
    elif type == 1:
        if not value:
            value = None
    elif type == 2:
        # NUMBER
        value_int = int(value)
        if value_int == value:
            value = value_int
        xf = book.xf_list[xf_index] # gets an XF object
        format_key = xf.format_key
        format = book.format_map[format_key] # gets a Format object
        format_str = format.format_str # this is the "number format string"
        if format_str.endswith(ur'\ "€"'):
            return (value, u'EUR')
        if format_str.endswith(ur'\ [$FRF]'):
            return (value, u'FRF')
        print value, format_str
#        TODO:
    elif type == 3:
        # DATE
        y, m, d, hh, mm, ss = xlrd.xldate_as_tuple(value, book.datemode)
        date = u'{0:04d}-{1:02d}-{2:02d}'.format(y, m, d) if any(n != 0 for n in (y, m, d)) else None
        value = u'T'.join(
            fragment
            for fragment in (
                date,
                u'{0:02d}:{1:02d}:{2:02d}'.format(hh, mm, ss)
                    if any(n != 0 for n in (hh, mm, ss)) or date is None
                    else None,
                )
            if fragment is not None
            )
    elif type == 4:
        value = bool(value)
    elif type == 5:
        # ERROR
        value = xlrd.error_text_from_code[value]
    return value



if __name__ == '__main__':
    bareme = 'Prestations'.encode('cp1252')
    path = ("P:/Legislation/Barèmes IPP/Barèmes IPP - ").encode('cp1252')
    handle_format_with_xlrd(path, bareme)

    # bleu foncé 49
    # bleu clair 27
    # blanc 64