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

Note:
    Currently this script requires an XLS version of the tables. XLSX file must be converted to XLS before use.
    To convert XLSX to XLS files, use Gnumeric, open XLSX file and save it as "MS Excel™ 97/2000/XP".

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

from biryani import baseconv, custom_conv, datetimeconv, states
from biryani import strings
import xlrd
import yaml


app_name = os.path.splitext(os.path.basename(__file__))[0]
conv = custom_conv(baseconv, datetimeconv, states)
french_date_re = re.compile(ur'(?P<day>0?[1-9]|[12]\d|3[01])/(?P<month>0?[1-9]|1[0-2])/(?P<year>[12]\d{3})$')
log = logging.getLogger(app_name)
N_ = lambda message: message
parameters = []
year_re = re.compile(ur'[12]\d{3}$')


class folded_unicode(unicode):
    pass


class literal_unicode(unicode):
    pass


class UnsortableList(list):
    def sort(self, *args, **kwargs):
        pass


class UnsortableOrderedDict(collections.OrderedDict):
    def items(self, *args, **kwargs):
        return UnsortableList(collections.OrderedDict.items(self, *args, **kwargs))


yaml.add_representer(folded_unicode, lambda dumper, data: dumper.represent_scalar(u'tag:yaml.org,2002:str',
    data, style='>'))
yaml.add_representer(literal_unicode, lambda dumper, data: dumper.represent_scalar(u'tag:yaml.org,2002:str',
    data, style='|'))
yaml.add_representer(unicode, lambda dumper, data: dumper.represent_scalar(u'tag:yaml.org,2002:str', data))
yaml.add_representer(UnsortableOrderedDict, yaml.representer.SafeRepresenter.represent_dict)


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


def get_hyperlink(sheet, row_index, column_index):
    return sheet.hyperlink_map.get((row_index, column_index))


def get_unmerged_cell_coordinates(row_index, column_index, merged_cells_tree):
    unmerged_cell_coordinates = merged_cells_tree.get(row_index, {}).get(column_index)
    if unmerged_cell_coordinates is None:
        return row_index, column_index
    return unmerged_cell_coordinates


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-s', '--source-dir', default = 'xls',
        help = 'path of source directory containing IPP XLS files')
    parser.add_argument('-t', '--target-dir', default = 'yaml', help = 'path of target directory for IPP JSON files')
    parser.add_argument('-v', '--verbose', action = 'store_true', default = False, help = "increase output verbosity")
    args = parser.parse_args()
    logging.basicConfig(level = logging.DEBUG if args.verbose else logging.WARNING, stream = sys.stdout)

    if not os.path.exists(args.target_dir):
        os.makedirs(args.target_dir)

    for filename in os.listdir(args.source_dir):
        if not filename.endswith('.xls'):
            continue
        log.info(u'Parsing file {}'.format(filename))
        book_name = os.path.splitext(filename)[0]
        xls_path = os.path.join(args.source_dir, filename).decode('utf-8')
        book = xlrd.open_workbook(filename = xls_path, formatting_info = True)

        book_yaml_dir = os.path.join(args.target_dir, book_name)
        if not os.path.exists(book_yaml_dir):
            os.makedirs(book_yaml_dir)

        sheet_names = [
            sheet_name
            for sheet_name in book.sheet_names()
            if not sheet_name.startswith((u'Abréviations', u'Outline'))
            ]
        sheet_title_by_name = {}
        for sheet_name in sheet_names:
            log.info(u'  Parsing sheet {}'.format(sheet_name))
            sheet = book.sheet_by_name(sheet_name)

            try:
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
                                    assert date_or_year.year < 2601, 'Invalid date {} in {} at row {}'.format(
                                        date_or_year, sheet_name, row_index + 1)
                                    values_rows.append(values_row)
                                    continue
                                if all(value in (None, u'') for value in values_row):
                                    # If first cell is empty and all other cells in line are also empty, ignore this
                                    # line.
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

                sheet_node = UnsortableOrderedDict()

                sheet_title = sheet_title_by_name.get(sheet_name)
                if sheet_title is not None:
                    sheet_node[u'Titre'] = sheet_title

                sheet_node[u'Titre court'] = sheet_name

                labels = []
                for labels_row in labels_rows:
                    for column_index, label in enumerate(labels_row):
                        if label is None:
                            continue
                        label = label.strip()
                        if not label:
                            continue
                        while column_index >= len(labels):
                            labels.append([])
                        column_labels = labels[column_index]
                        if not column_labels or column_labels[-1] != label:
                            column_labels.append(label)
                labels = [
                    tuple(column_labels1) if column_labels1 else (u'Colonne {} sans titre',)
                    for index, column_labels1 in enumerate(labels, 1)
                    ]

                taxipp_name_by_column_labels = UnsortableOrderedDict()
                for column_labels, taxipp_name in zip(labels, taxipp_names_row):
                    if not taxipp_name:
                        continue
                    taxipp_name_by_column_label = taxipp_name_by_column_labels
                    for column_label in column_labels[:-1]:
                        taxipp_name_by_column_label = taxipp_name_by_column_label.setdefault(column_label,
                            UnsortableOrderedDict())
                    taxipp_name_by_column_label[column_labels[-1]] = taxipp_name
                if taxipp_name_by_column_labels:
                    sheet_node[u'Noms TaxIPP'] = taxipp_name_by_column_labels

                sheet_values = []
                for value_row in values_rows:
                    cell_by_column_labels = UnsortableOrderedDict()
                    for column_labels, cell in zip(labels, value_row):
                        if cell is None or cell == '':
                            continue
                        cell_by_column_label = cell_by_column_labels
                        for column_label in column_labels[:-1]:
                            cell_by_column_label = cell_by_column_label.setdefault(column_label,
                                UnsortableOrderedDict())
                        # Merge (amount, unit) couples to a string to simplify YAML.
                        cell_by_column_label[column_labels[-1]] = u' '.join(
                            unicode(fragment)
                            for fragment in cell
                            ) if isinstance(cell, tuple) else cell
                    sheet_values.append(cell_by_column_labels)
                if sheet_values:
                    sheet_node[u'Valeurs'] = sheet_values

                notes_lines = [
                    u' | '.join(
                        cell for cell in row
                        if cell
                        )
                    for row in notes_rows
                    ]
                if notes_lines:
                    sheet_node[u'Notes'] = literal_unicode(u'\n'.join(notes_lines))

                descriptions_lines = [
                    u' | '.join(
                        cell for cell in row
                        if cell
                        )
                    for row in descriptions_rows
                    ]
                if descriptions_lines:
                    sheet_node[u'Description'] = literal_unicode(u'\n'.join(descriptions_lines))

                with open(os.path.join(book_yaml_dir, strings.slugify(sheet_name) + '.yaml'), 'w') as yaml_file:
                    yaml.dump(sheet_node, yaml_file, allow_unicode = True, default_flow_style = False, indent = 2,
                        width = 120)
            except:
                log.exception(u'An exception occurred when parsing sheet "{}" of XLS file "{}"'.format(sheet_name,
                    filename))

    return 0


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
                u'#,##0',
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
    if isinstance(cell, tuple):
        # Replace (value, unit) couple to a string.
        cell = u' '.join(
            unicode(fragment)
            for fragment in cell
            )
    assert cell is None or isinstance(cell, basestring), u'Expected a string. Got: {}'.format(cell).encode('utf-8')
    return cell


if __name__ == "__main__":
    sys.exit(main())
