import os

import pandas as pd
#import openpyxl
from pathlib import Path


def get_excel_files():
    # Get all excel files in directory
    p = Path('.')
    return p.glob('*.xlsx')


field_renames = {
    #'אני מעוניין להרשם ל-': 'workshop',
    #'מעוניינים להשתתף ב:' : 'workshop',
    #'אני מעוניין/נת להרשם ל-': 'workshop',
    'שם פרטי ומשפחה - נער/ה': 'name',
    'שם פרטי ומשפחה': 'name',
    'עיר מגורים': 'city',
    'מספר משתתפים': 'participantNum',
    'איך שמעת עלינו': 'foundUsVia',
    'היכן שמעתם על הסדנה?': 'foundUsVia',
    'שולם': 'amountPaid'
}

default_vals = {
    'amountPaid': 0,
    'participantNum': 1,
    'foundUsVia': 'לא ידוע',
    'city': 'לא ידוע'
}


def list_sheets(filename: str):
    # List sheets in excel file
    xls = pd.ExcelFile(filename)
    sheet_dict = pd.read_excel(xls, sheet_name=None)
    return sheet_dict


def read_excel(sheet):
    # Read Excel file
    df = pd.read_excel('data.xlsx', sheet_name=sheet)
    return df


'''
def consolidate(sheets):
    # Consolidate all sheets into one dataframe
    df = pd.concat([read_excel(sheet) for sheet in sheets])
    return df
'''


def print_fields(sheet_dict):
    for sheet_entry in sheet_dict.items():
        print(sheet_entry[0], end=': ')
        print(sheet_entry[1].columns)


def resturcture_sheet(filename: str, workshop: str, sheet: pd.DataFrame):
    column_names = sheet.columns
    rename_keys = list(field_renames.keys())
    for column_name in column_names:
        if column_name in rename_keys:
            sheet.rename(columns={column_name: field_renames[column_name]}, inplace=True)
        else:
            sheet.drop(column_name, axis=1, inplace=True)
    for field in default_vals.keys():
        if field not in sheet.columns:
            sheet[field] = default_vals[field]
    sheet['workshop'] = workshop
    sheet['filename'] = filename


def main():
    empty = {
        'workshop': [],
        'name': [],
        'city': [],
        'participantNum': [],
        'foundUsVia': [],
        'amountPaid': []
    }
    consolidated_sheet = pd.DataFrame()

    for file in get_excel_files():
        filename = file.name
        if os.path.basename(filename) == 'consolidated.xlsx':
            continue
        sheet_dict = list_sheets(filename)
        for sheet_pair in sheet_dict.items():
            resturcture_sheet(file.stem, sheet_pair[0], sheet_pair[1])

        to_concat = list(sheet_dict.values())
        to_concat.append(consolidated_sheet)
        consolidated_sheet = pd.concat(to_concat)

    consolidated_sheet.to_excel('consolidated.xlsx', index=False)


if __name__ == '__main__':
    main()


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
