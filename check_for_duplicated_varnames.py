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


import argparse
import datetime
import math
import os

import numpy as np
import pandas as pd
# Architecture :
# un xlsx contient des sheets qui contiennent des variables, chaque sheet ayant un vecteur de dates


def clean_date(date_time):
    ''' Conversion des dates spécifiées en année au format year/01/01
    Remise des jours au premier du mois '''
    if len(str(date_time)) == 4 :
        return datetime.date(date_time, 1, 1)
    else:
        return date_time.date().replace(day = 1)


def clean_sheet(xls_file, sheet_name):
    ''' Cleaning excel sheets and creating small database'''

    sheet = xls_file.parse(sheet_name, index_col = None)

    # Conserver les bonnes colonnes : on drop tous les "Unnamed"
    for col in sheet.columns.values:
        if col[0:7] == 'Unnamed':
            sheet = sheet.drop([col], 1)

    # Pour l'instant on drop également tous les ref_leg, jorf et notes
    for var_to_drop in ['ref_leg', 'jorf', 'Notes', 'notes', 'date_ir'] :
        if var_to_drop in sheet.columns.values:
            sheet = sheet.drop(var_to_drop, axis = 1)


    # Pour impôt sur le revenu, il y a date_IR et date_rev : on utilise date_rev, que l'on renome date pour plus de cohérence
    if 'date_rev' in sheet.columns.values:
            sheet = sheet.rename(columns={'date_rev':u'date'})

    # Conserver les bonnes lignes : on drop s'il y a du texte ou du NaN dans la colonne des dates
    def is_var_nan(row,col):
        return isinstance(sheet.iloc[row, col], float) and math.isnan(sheet.iloc[row, col])

    sheet['date_absente'] = False
    for i in range(0,sheet.shape[0]):
        sheet.loc[i,['date_absente']] = isinstance(sheet.iat[i,0], basestring) or is_var_nan(i,0)
    sheet = sheet[sheet.date_absente == False]
    sheet = sheet.drop(['date_absente'], axis = 1)

    # S'il y a du texte au milieu du tableau (explications par exemple) => on le transforme en NaN
    for col in range(0, sheet.shape[1]):
        for row in range(0,sheet.shape[0]):
            if isinstance(sheet.iloc[row,col], unicode):
                sheet.iat[row,col] = np.nan

    # Gérer la suppression et la création progressive de dispositifs

    sheet.iloc[0, :].fillna('-', inplace = True)

    assert 'date' in sheet.columns, "Aucune colonne date dans la feuille : {}".format(sheet)
    sheet['date'] =[ clean_date(d) for d in  sheet['date']]

    return sheet

def dic_of_same_variable_names(xls_file, sheet_names):
    dic = {}
    all_variables = np.zeros(1)
    multiple_names = []
    for sheet_name in  sheet_names:
        dic[sheet_name]= clean_sheet(xls_file, sheet_name)
        sheet = clean_sheet(xls_file, sheet_name)
        columns =  np.delete(sheet.columns.values,0)
        all_variables = np.append(all_variables,columns)
    for i in range(0,len(all_variables)):
        var = all_variables[i]
        new_variables = np.delete(all_variables,i)
        if var in new_variables:
            multiple_names.append(str(var))
    multiple_names = list(set(multiple_names))
    dic_var_to_sheet={}
    for sheet_name in sheet_names:
        sheet = clean_sheet(xls_file, sheet_name)
        columns =  np.delete(sheet.columns.values,0)
        for var in multiple_names:
            if var in columns:
                if var in dic_var_to_sheet.keys():
                    dic_var_to_sheet[var].append(sheet_name)
                else:
                    dic_var_to_sheet[var] = [sheet_name]
    return dic_var_to_sheet


if __name__ == '__main__':
    path = u"P:/Legislation/Barèmes IPP/"
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--dir', default = path, help = 'path of IPP XLS directory')
    args = parser.parse_args()

    baremes = [u'Prestations', u'prélèvements sociaux', u'Impôt Revenu']
    forbiden_sheets = {u'Impôt Revenu' : (u'Barème IGR',),
                       u'prélèvements sociaux' : (u'Abréviations', u'ASSIETTE PU', u'AUBRYI')}
    for bareme in baremes :
        xls_path = os.path.join(args.dir, u"Barèmes IPP - {0}.xlsx".format(bareme))
        xls_file = pd.ExcelFile(xls_path)

        # Retrait des onglets qu'on ne souhaite pas importer
        sheets_to_remove = (u'Sommaire', u'Outline')
        if bareme in forbiden_sheets.keys():
            sheets_to_remove += forbiden_sheets[bareme]

        sheet_names = [
            sheet_name
            for sheet_name in xls_file.sheet_names
            if not sheet_name.startswith(sheets_to_remove)
            ]

        # Test si deux variables ont le même nom
        test_duplicate = dic_of_same_variable_names(xls_file, sheet_names)
        assert not test_duplicate, u'Au moins deux variables ont le même nom dans le classeur {} : u{}'.format(
            bareme,test_duplicate)
