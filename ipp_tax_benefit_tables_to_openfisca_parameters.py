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


"""Extract parameters from IPP's tax benefit tables.

Note: Currently this script requires an XLS version of the tables. XLSX file must be converted to XLS before use.

IPP = Institut des politiques publiques
http://www.ipp.eu/en/tools/ipp-tax-and-benefit-tables/
http://www.ipp.eu/fr/outils/baremes-ipp/
"""


import argparse
import collections
import datetime
import itertools
import logging
import os
import re
import sys
import textwrap

from biryani import baseconv, custom_conv, datetimeconv, states
from biryani import strings
import xlrd


app_name = os.path.splitext(os.path.basename(__file__))[0]
baremes = [
    # u'Chomage',
    # u'Impot Revenu',
    # u'Marche du travail',
    u'prelevements sociaux',
    # u'Prestations',
    # u'Taxation indirecte',
    # u'Taxation du capital',
    # u'Taxes locales',
    ]
conv = custom_conv(baseconv, datetimeconv, states)
forbiden_sheets = {
    # u'Impot Revenu': (u'Barème IGR',),
    u'prelevements sociaux': (
        u'ASSIETTE PU',
        u'AUBRYI',
        # u'AUBRYII',
        u'CNRACL',
        u'FILLON',
        ),
    # u'Taxation indirecte': (u'TVA par produit',),
    }
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


cell_to_date = conv.condition(
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


# currency_converter = conv.first_match(
#     conv.pipe(
#         conv.test_isinstance(basestring),
#         conv.cleanup_line,
#         conv.test_none(),
#         ),
#     conv.pipe(
#         conv.test_isinstance(tuple),
#         conv.test(lambda couple: len(couple) == 2, error = N_(u"Invalid couple length")),
#         conv.struct(
#             (
#                 conv.pipe(
#                     conv.test_isinstance((float, int)),
#                     conv.not_none,
#                     ),
#                 conv.pipe(
#                     conv.test_isinstance(basestring),
#                     conv.test_in([
#                         u'%',
#                         u'EUR',
#                         u'FRF',
#                         ]),
#                     ),
#                 ),
#             ),
#         ),
#     )


currency_or_number_converter = conv.first_match(
    conv.test_isinstance(float),
    conv.test_isinstance(int),
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
                        u'%',
                        u'EUR',
                        u'FRF',
                        ]),
                    ),
                ),
            ),
        ),
    )


def rename_keys(new_key_by_old_key):
    def rename_keys_converter(value, state = None):
        if value is None:
            return value, None
        renamed_value = value.__class__()
        for item_key, item_value in value.iteritems():
            renamed_value[new_key_by_old_key.get(item_key, item_key)] = item_value
        return renamed_value, None

    return rename_keys_converter


values_row_converter = conv.pipe(
    rename_keys({
        u"Date d'effet": u"Date d'entrée en vigueur",
        u"Note": u"Notes",
        u"Publication au JO": u"Parution au JO",
        u"Publication  JO": u"Parution au JO",
        u"Publication JO": u"Parution au JO",
        u"Référence": u"Références législatives",
        u"Référence législative": u"Références législatives",
        u"Références législatives                  (taux d'appel)": u"Références législatives",
        u"Références législatives                  (taux de cotisation)": u"Références législatives",
        u"Références législatives ou BOI": u"Références législatives",
        u"Remarques": u"Notes",
        }),
    conv.struct(
        collections.OrderedDict((
            (u"Date d'entrée en vigueur", conv.pipe(
                conv.test_isinstance(basestring),
                conv.iso8601_input_to_date,
                conv.not_none,
                )),
            (u"Références législatives", conv.pipe(
                conv.test_isinstance(basestring),
                conv.cleanup_line,
                )),
            (u"Parution au JO", conv.pipe(
                conv.test_isinstance(basestring),
                conv.iso8601_input_to_date,
                conv.date_to_iso8601_str,
                )),
            (u"Notes", conv.pipe(
                conv.test_isinstance(basestring),
                conv.cleanup_line,
                )),
            (None, conv.pipe(
                conv.test_isinstance(basestring),
                conv.cleanup_line,
                conv.test_none(),
                )),
            )),
        default = currency_or_number_converter,
        ),
    )


def escape_xml(value):
    if value is None:
        return value
    if isinstance(value, str):
        return value.decode('utf-8')
    if not isinstance(value, unicode):
        value = unicode(value)
    return value.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def get_hyperlink(sheet, row_index, column_index):
    return sheet.hyperlink_map.get((row_index, column_index))


def get_unmerged_cell_coordinates(row_index, column_index, merged_cells_tree):
    unmerged_cell_coordinates = merged_cells_tree.get(row_index, {}).get(column_index)
    if unmerged_cell_coordinates is None:
        return row_index, column_index
    return unmerged_cell_coordinates


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--dir', default = 'Baremes_IPP_2015', help = 'path of IPP XLS directory')
    parser.add_argument('-v', '--verbose', action = 'store_true', default = False, help = "increase output verbosity")
    args = parser.parse_args()
    # args.dir = path
    logging.basicConfig(level = logging.DEBUG if args.verbose else logging.WARNING, stream = sys.stdout)

    root_node = dict(
        children = [],
        name = "root",
        text = textwrap.dedent(u"""\
            Ce document présente l'ensemble de la législation permettant le calcul des contributions sociales, taxes sur
            les salaires  et cotisations sociales. Il s'agit des barèmes bruts de la législation utilisés dans le
            micro-simulateur de l'IPP, TAXIPP. Les sources législatives (texte de loi, numéro du décret ou arrêté) ainsi
            que la date de publication au Journal Officiel de la République française (JORF) sont systématiquement
            indiquées. La première ligne du fichier (masquée) indique le nom des paramètres dans TAXIPP.

            Citer cette source :
            Barèmes IPP: prélèvements sociaux, Institut des politiques publiques, avril 2014.

            Auteurs :
            Antoine Bozio, Julien Grenet, Malka Guillot, Laura Khoury et Marianne Tenand

            Contacts :
            marianne.tenand@ipp.eu; antoine.bozio@ipp.eu; malka.guillot@ipp.eu

            Licence :
            Licence ouverte / Open Licence
            """).split(u'\n'),
        title = u"Barème IPP",
        type = u'NODE',
        )

    for bareme in baremes:
        xls_path = os.path.join(args.dir.decode('utf-8'), u"Baremes IPP - {0}.xls".format(bareme))
        if not os.path.exists(xls_path):
            log.warning("Skipping file {} that doesn't exist: {}".format(bareme, xls_path))
            continue
        log.info(u'Parsing file {}'.format(bareme))
        book = xlrd.open_workbook(filename = xls_path, formatting_info = True)

        sheet_names = [
            sheet_name
            for sheet_name in book.sheet_names()
            if not sheet_name.startswith((u'Abréviations', u'Outline')) and sheet_name not in forbiden_sheets.get(
                bareme, [])
            ]
        sheet_title_by_name = {}
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

            if sheet_name.startswith(u'Sommaire'):
                # Associate the titles of the sheets to their Excel names.
                for row_index in range(sheet.nrows):
                    linked_sheet_number = transform_xls_cell_to_json(book, sheet, merged_cells_tree, row_index, 2)
                    if isinstance(linked_sheet_number, int):
                        linked_sheet_title = transform_xls_cell_to_str(book, sheet, merged_cells_tree, row_index, 3)
                        if linked_sheet_title is not None:
                            hyperlink = get_hyperlink(sheet, row_index, 3)
                            if hyperlink is not None and hyperlink.type == u'workbook':
                                linked_sheet_name = hyperlink.textmark.split(u'!', 1)[0].strip(u'"').strip(u"'")
                                sheet_title_by_name[linked_sheet_name] = linked_sheet_title
                continue

            descriptions_rows = []
            labels_rows = []
            notes_rows = []
            state = 'taxipp_names'
            taxipp_names_row = None
            values_rows = []
            for row_index in range(sheet.nrows):
                columns_count = len(sheet.row_values(row_index))
                if state == 'taxipp_names':
                    taxipp_names_row = [
                        taxipp_name
                        for taxipp_name in (
                            transform_xls_cell_to_str(book, sheet, merged_cells_tree, row_index, column_index)
                            for column_index in range(columns_count)
                            )
                        ]
                    state = 'labels'
                    continue
                if state == 'labels':
                    first_cell_value = transform_xls_cell_to_json(book, sheet, merged_cells_tree, row_index, 0)
                    date_or_year, error = conv.pipe(
                        conv.test_isinstance((int, basestring)),
                        cell_to_date,
                        conv.not_none,
                        )(first_cell_value, state = conv.default_state)
                    if error is not None:
                        # First cell of row is not a date => Assume it is a label.
                        labels_rows.append([
                            transform_xls_cell_to_str(book, sheet, merged_cells_tree, row_index, column_index)
                            for column_index in range(columns_count)
                            ])
                        continue
                    state = 'values'
                if state == 'values':
                    first_cell_value = transform_xls_cell_to_json(book, sheet, merged_cells_tree, row_index, 0)
                    if first_cell_value is None or isinstance(first_cell_value, (int, basestring)):
                        date_or_year, error = cell_to_date(first_cell_value, state = conv.default_state)
                        if error is None:
                            # First cell of row is a valid date or year.
                            values_row = [
                                transform_xls_cell_to_json(book, sheet, merged_cells_tree, row_index, column_index)
                                for column_index in range(columns_count)
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
                            for column_index in range(columns_count)
                            ])
                        continue
                    state = 'description'
                assert state == 'description'
                descriptions_rows.append([
                    transform_xls_cell_to_str(book, sheet, merged_cells_tree, row_index, column_index)
                    for column_index in range(columns_count)
                    ])

            text_lines = []
            for row in notes_rows:
                text_lines.append(u' | '.join(
                    cell for cell in row
                    if cell
                    ))
            if text_lines:
                text_lines.append(None)
            for row in descriptions_rows:
                text_lines.append(u' | '.join(
                    cell for cell in row
                    if cell
                    ))

            sheet_title = sheet_title_by_name.get(sheet_name)
            if sheet_title is None:
                log.warning(u"Missing title for sheet {} in summary".format(sheet_name))
                continue
            labels = []
            for labels_row in labels_rows:
                for column_index, label in enumerate(labels_row):
                    if not label:
                        continue
                    while column_index >= len(labels):
                        labels.append([])
                    labels_column = labels[column_index]
                    if not labels_column or labels_column[-1] != label:
                        labels_column.append(label)
            labels = [
                tuple(labels_column1) if len(labels_column1) > 1 else labels_column1[0]
                for labels_column1 in labels
                ]

            cell_by_label_rows = []
            for value_row in values_rows:
                cell_by_label = collections.OrderedDict(itertools.izip(labels, value_row))
                cell_by_label, errors = values_row_converter(cell_by_label, state = conv.default_state)
                assert errors is None, "Errors in {}:\n{}".format(cell_by_label, errors)
                cell_by_label_rows.append(cell_by_label)

            sheet_node = dict(
                children = [],
                name = strings.slugify(sheet_name, separator = u'_'),
                text = text_lines,
                title = sheet_title,
                type = u'NODE',
                )
            root_node['children'].append(sheet_node)

            for taxipp_name, labels_column in zip(taxipp_names_row, labels):
                if not taxipp_name or taxipp_name in (u'date',):
                    continue
                variable_node = dict(
                    children = [],
                    name = strings.slugify(taxipp_name, separator = u'_'),
                    title = u' - '.join(labels_column) if isinstance(labels_column, tuple) else labels_column,
                    type = u'CODE',
                    )
                sheet_node['children'].append(variable_node)

                for cell_by_label in cell_by_label_rows:
                    amount_and_unit = cell_by_label[labels_column]
                    variable_node['children'].append(dict(
                        law_reference = cell_by_label[u'Références législatives'],
                        notes = cell_by_label[u'Notes'],
                        publication_date = cell_by_label[u"Parution au JO"],
                        start_date = cell_by_label[u"Date d'entrée en vigueur"],
                        type = u'VALUE',
                        unit = amount_and_unit[1] if isinstance(amount_and_unit, tuple) else None,
                        value = amount_and_unit[0] if isinstance(amount_and_unit, tuple) else amount_and_unit,
                        ))

            # dates = [
            #     conv.check(cell_to_date)(
            #         row[1] if bareme == u'Impot Revenu' else row[0],
            #         state = conv.default_state,
            #         )
            #     for row in values_rows
            #     ]
            # for column_index, taxipp_name in enumerate(taxipp_names_row):
            #     if taxipp_name and strings.slugify(taxipp_name) not in (
            #             'date',
            #             'date-ir',
            #             'date-rev',
            #             'note',
            #             'notes',
            #             'ref-leg',
            #             ):
            #         vector = [
            #             transform_cell_value(date, row[column_index])
            #             for date, row in zip(dates, values_rows)
            #             ]
            #         vector = [
            #             cell if not isinstance(cell, basestring) or cell == u'nc' else '-'
            #             for cell in vector
            #             ]
            #         # vector_by_taxipp_name[taxipp_name] = pd.Series(vector, index = dates)
            #         vector_by_taxipp_name[taxipp_name] = vector
            #

    print_node(root_node)

    return 0


def print_node(node, indent = 0):
    attributes = node.copy()
    children = attributes.pop('children', None)
    text = attributes.pop('text', None)
    if text:
        while text and not (text[0] and text[0].strip()):
            del text[0]
        while text and not (text[-1] and text[-1].strip()):
            del text[-1]
    type = attributes.pop('type')
    print u'{}<{}{}{}>'.format(
        u'  ' * indent,
        type,
        u''.join(
            u' {}="{}"'.format(name, escape_xml(value))
            for name, value in sorted(attributes.iteritems())
            if value is not None
            ),
        u'' if children or text else u'/',
        ).encode('utf-8')
    if text:
        for line in text:
            if line and line.strip():
                print u'{}{}'.format(u'  ' * (indent + 1), escape_xml(line)).encode('utf-8')
            else:
                print
    if children or text:
        for child in children:
            print_node(child, indent = indent + 1)
        print u'{}</{}>'.format(u'  ' * indent, type).encode('utf-8')


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
                (u'{0:02d}:{1:02d}:{2:02d}'.format(hh, mm, ss)
                    if any(n != 0 for n in (hh, mm, ss)) or date is None
                    else None),
                )
            if fragment is not None
            )
    elif type == 4:
        value = bool(value)
    elif type == 5:
        # ERROR
        value = xlrd.error_text_from_code[value]
    # elif type == 6:
    #     TODO
    # else:
    #     assert False, str((type, value))
    return value


def transform_xls_cell_to_str(book, sheet, merged_cells_tree, row_index, column_index):
    cell = transform_xls_cell_to_json(book, sheet, merged_cells_tree, row_index, column_index)
    assert cell is None or isinstance(cell, basestring), u'Expected a string. Got: {}'.format(cell).encode('utf-8')
    return cell


if __name__ == "__main__":
    sys.exit(main())
