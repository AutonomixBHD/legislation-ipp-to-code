#! /usr/bin/env python
# -*- coding: utf-8 -*-


# Law-to-Code -- Extract formulas & parameters from laws
# By: Emmanuel Raviart <emmanuel@raviart.com>
#
# Copyright (C) 2013, 2014 OpenFisca Team
# https://github.com/openfisca/LawToCode
#
# This file is part of Law-to-Code.
#
# Law-to-Code is free software; you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as
# published by the Free Software Foundation, either version 3 of the
# License, or (at your option) any later version.
#
# Law-to-Code is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.


"""Extract parameters from IPP's tax benefit tables.

Note: Currently this script requires an XLS version of the tables. XLSX file must be converted to XLS before use.

IPP = Institut des politiques publiques
http://www.ipp.eu/en/tools/ipp-tax-and-benefit-tables/
http://www.ipp.eu/fr/outils/baremes-ipp/
"""


import argparse
import collections
import datetime
import logging
import os
import re
import sys

from biryani1 import baseconv, custom_conv, datetimeconv, states
from biryani1 import strings
import numpy as np
import pandas as pd
import xlrd

app_name = os.path.splitext(os.path.basename(__file__))[0]
conv = custom_conv(baseconv, datetimeconv, states)
french_date_re = re.compile(ur'(?P<day>0?[1-9]|[12]\d|3[01])/(?P<month>0?[1-9]|1[0-2])/(?P<year>[12]\d{3})$')
log = logging.getLogger(app_name)
N_ = lambda message: message
parameters = []
year_re = re.compile(ur'[12]\d{3}$')


def input_to_french_date(value, state = None):
    if value is None:
        return None, None
    if state is None:
        state = conv.default_state
    match = french_date_re.match(value)
    if match is None:
        return value, state._(u'Invalid french date')
    return datetime.date(int(match.group('year')), int(match.group('month')), int(match.group('day'))), None


cell_to_date_or_year = conv.condition(
    conv.test_isinstance(int),
    conv.pipe(
        conv.test_between(1914, 2020),
        conv.function(lambda year: datetime.date(year, 1, 1)),
        ),
    conv.pipe(
        conv.test_isinstance(basestring),
        conv.first_match(
            conv.pipe(
                conv.test(lambda date: year_re.match(date), error = 'Not a valid year'),
                conv.function(lambda year: datetime.date(year, 1, 1)),
                ),
            input_to_french_date,
            conv.iso8601_input_to_date,
            ),
        ),
    )


currency_converter = conv.first_match(
    conv.pipe(
        conv.test_isinstance(basestring),
        conv.cleanup_line,
        conv.test_none(),
        ),
    conv.pipe(
        conv.test_isinstance(tuple),
        conv.test(lambda couple: len(couple) == 2, error = N_(u"Invalid couple length")),
        conv.struct(
            (
                conv.pipe(
                    conv.test_isinstance((float, int)),
                    conv.not_none,
                    ),
                conv.pipe(
                    conv.test_isinstance(basestring),
                    conv.test_in([
                        u'EUR',
                        u'FRF',
                        ]),
                    ),
                ),
            ),
        ),
    )


pss_converters = collections.OrderedDict((
    (u"Date d'effet", conv.pipe(
        conv.test_isinstance(basestring),
        conv.iso8601_input_to_date,
        conv.date_to_iso8601_str,
        conv.not_none,
        )),
    (u'Plafond de la Sécurité sociale (mensuel)', currency_converter),
    (u'Plafond de la Sécurité sociale (annuel)', currency_converter),
    (u'Référence législative', conv.pipe(
        conv.test_isinstance(basestring),
        conv.cleanup_line,
        )),
    (u'Parution au JO', conv.pipe(
        conv.test_isinstance(basestring),
        conv.iso8601_input_to_date,
        conv.date_to_iso8601_str,
        )),
    (u'Notes', conv.pipe(
        conv.test_isinstance(basestring),
        conv.cleanup_line,
        )),
    (None, conv.pipe(
        conv.test_isinstance(basestring),
        conv.cleanup_line,
        conv.test_none(),
        )),
    ))


def get_unmerged_cell_coordinates(row_index, column_index, merged_cells_tree):
    unmerged_cell_coordinates = merged_cells_tree.get(row_index, {}).get(column_index)
    if unmerged_cell_coordinates is None:
        return row_index, column_index
    return unmerged_cell_coordinates


def main(path, date):
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--dir', default = path + date, help = 'path of IPP XLS directory')
    parser.add_argument('-v', '--verbose', action = 'store_true', default = False, help = "increase output verbosity")
    args = parser.parse_args()
    #args.dir = path
    logging.basicConfig(level = logging.DEBUG if args.verbose else logging.WARNING, stream = sys.stdout)

    forbiden_sheets = {
        u'Impot Revenu': (u'Barème IGR',),
        u'prelevements sociaux': (u'Abréviations', u'ASSIETTE PU', u'AUBRYI',  u'AUBRYII'),
        u'Taxation indirecte': (u'TVA par produit',),
        }
    baremes = [u'Chomage', u'Impot Revenu', u'prelevements sociaux', u'Prestations',u'Taxation indirecte',u'Taxation du capital',u'Taxes locales',u'Marche du travail',]
#    baremes_TODO = [u'Taxation du capital', u'Impôt Revenu', u'Marché du travail', u'Chômage', u'Retraite', u'Taxes locales', ]
    for bareme in baremes:
        log.info(u'Parsing file {}'.format(bareme))
        xls_path = os.path.join(args.dir.decode('utf-8'), u"Baremes IPP - {0}.xls".format(bareme))
#       xls_path = os.path.join(path, u"Baremes IPP - {0}.xls".format(bareme))
        book = xlrd.open_workbook(filename = xls_path, formatting_info = True)
        sheet_names = [
            sheet_name
            for sheet_name in book.sheet_names()
            if not sheet_name.startswith((u'Sommaire', u'Outline'))
                and not sheet_name in forbiden_sheets.get(bareme, [])
            ]
        vector_by_taxipp_name = {}
        for sheet_name in sheet_names:
            log.info(u'  Parsing sheet {}'.format(sheet_name))
            sheet = book.sheet_by_name(sheet_name)

            # Extract coordinates of merged cells.
            merged_cells_tree = {}
            for row_low, row_high, column_low, column_high in sheet.merged_cells:
                for row_index in range(row_low, row_high):
                    cell_coordinates_by_merged_column_index = merged_cells_tree.setdefault(
                        row_index, {})
                    for column_index in range(column_low, column_high):
                        cell_coordinates_by_merged_column_index[column_index] = (row_low, column_low)

            descriptions_rows = []
            labels_rows = []
            notes_rows = []
            state = 'taxipp_names'
            taxipp_names_row = None
            values_rows = []
            for row_index in range(sheet.nrows):
                ncols = len(sheet.row_values(row_index))
                if state == 'taxipp_names':
                    taxipp_names_row = [
                        taxipp_name
                        for taxipp_name in (
                            transform_xls_cell_to_str(book, sheet, merged_cells_tree, row_index, column_index)
                            for column_index in range(ncols)
                            )
                        ]
                    state = 'labels'
                    continue
                if state == 'labels':
                    first_cell_value = transform_xls_cell_to_json(book, sheet, merged_cells_tree, row_index, 0)
                    date_or_year, error = conv.pipe(
                        conv.test_isinstance((int, basestring)),
                        cell_to_date_or_year,
                        conv.not_none,
                        )(first_cell_value, state = conv.default_state)
                    if error is not None:
                        # First cell of row is not a date => Assume it is a label.
                        labels_rows.append([
                            transform_xls_cell_to_str(book, sheet, merged_cells_tree, row_index, column_index)
                            for column_index in range(ncols)
                            ])
                        continue
                    state = 'values'
                if state == 'values':
                    first_cell_value = transform_xls_cell_to_json(book, sheet, merged_cells_tree, row_index, 0)
                    if first_cell_value is None or isinstance(first_cell_value, (int, basestring)):
                        date_or_year, error = cell_to_date_or_year(first_cell_value, state = conv.default_state)
                        if error is None:
                            # First cell of row is a valid date or year.
                            values_row = [
                                transform_xls_cell_to_json(book, sheet, merged_cells_tree, row_index, column_index)
                                for column_index in range(ncols)
                                ]
                            if date_or_year is not None:
                                assert date_or_year.year < 2601, 'Invalid date {} in {} at row {}'.format(date_or_year,
                                    sheet_name, row_index + 1)
                                values_rows.append(values_row)
                                continue
                            if all(value in (None, u'') for value in values_row):
                                # If first cell is empty and all other cells in line are also empty, ignore this line.
                                continue
                            # First cell has no date and other cells in row are not empty => Assume it is a note.
                    state = 'notes'
                if state == 'notes':
                    first_cell_value = transform_xls_cell_to_json(book, sheet, merged_cells_tree, row_index, 0)
                    if isinstance(first_cell_value, basestring) and first_cell_value.strip().lower() == 'notes':
                        notes_rows.append([
                            transform_xls_cell_to_str(book, sheet, merged_cells_tree, row_index, column_index)
                            for column_index in range(ncols)
                            ])
                        continue
                    state = 'description'
                assert state == 'description'
                descriptions_rows.append([
                    transform_xls_cell_to_str(book, sheet, merged_cells_tree, row_index, column_index)
                    for column_index in range(ncols)
                    ])

            dates = [
                conv.check(cell_to_date_or_year)(
                    row[1] if bareme == u'Impot Revenu' else row[0],
                    state = conv.default_state,
                    ).replace(day = 1)
                for row in values_rows
                ]
            for column_index, taxipp_name in enumerate(taxipp_names_row):
                if taxipp_name and strings.slugify(taxipp_name) not in ('date', 'date-ir', 'date-rev', 'note', 'ref-leg', 'notes') :
                    vector = [
                        transform_cell_value(date, row[column_index])
                        for date, row in zip(dates, values_rows)
                        ]
                    vector = [
                        cell if not isinstance(cell, basestring) else np.nan
                        for cell in vector
                        ]
                    vector_by_taxipp_name[taxipp_name] = pd.Series(vector, index = dates)
        months = [
            datetime.date(year, month, 1)
            for year in range(1914, 2021)
            for month in range(1, 13)
            ]
        data_frame = pd.DataFrame(index = months)
        for taxipp_name, vector in vector_by_taxipp_name.iteritems():
            data_frame[taxipp_name] = np.nan
            data_frame.loc[vector.index.values, taxipp_name] = vector.values
        data_frame.fillna(method = 'pad', inplace = True)
        data_frame.dropna(axis = 0, how = 'all', inplace = True)
        data_frame.to_csv(args.dir + bareme + '.csv', encoding = 'utf-8')
        print u"Voilà, la table agrégée de {} est créée !".format(bareme)

    return 0


def transform_cell_value(date, cell_value):
    if isinstance(cell_value, tuple):
        value, currency = cell_value
        if currency == u'FRF':
            if date < datetime.date(1960, 1, 1):
                return round(value / (100 * 6.55957), 2)
            return round(value / 6.55957, 2)
        return value
    return cell_value


def transform_xls_cell_to_json(book, sheet, merged_cells_tree, row_index, column_index):
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
    unmerged_cell_coordinates = merged_cells_tree.get(row_index, {}).get(column_index)
    if unmerged_cell_coordinates is None:
        unmerged_row_index = row_index
        unmerged_column_index = column_index
    else:
        unmerged_row_index, unmerged_column_index = unmerged_cell_coordinates
    type = sheet.row_types(unmerged_row_index)[unmerged_column_index]
    value = sheet.row_values(unmerged_row_index)[unmerged_column_index]
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
        xf_index = sheet.cell_xf_index(row_index, column_index)
        xf = book.xf_list[xf_index]  # Get an XF object.
        format_key = xf.format_key
        format = book.format_map[format_key]  # Get a Format object.
        format_str = format.format_str  # This is the "number format string".
        if format_str in (
                u'0',
                u'General',
                u'GENERAL',
                u'_-* #,##0\ _€_-;\-* #,##0\ _€_-;_-* \-??\ _€_-;_-@_-',
                ) or format_str.endswith(u'0.00'):
            return value
        if u'€' in format_str:
            return (value, u'EUR')
        if u'FRF' in format_str or ur'\F\R\F' in format_str:
            return (value, u'FRF')
        assert format_str.endswith(u'%'), 'Unexpected format "{}" for value: {}'.format(format_str, value)
        return (value, u'%')
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


def transform_xls_cell_to_str(book, sheet, merged_cells_tree, row_index, column_index):
    cell = transform_xls_cell_to_json(book, sheet, merged_cells_tree, row_index, column_index)
    assert cell is None or isinstance(cell, basestring), u'Expected a string. Got: {}'.format(cell).encode('utf-8')
    return cell


if __name__ == "__main__":
    path = 'Directory of Baremes' 
    sys.exit(main(path, date = "28_04")) # date = quantième et numéro du mois (répertoire des fichiers .xls barèmes)
