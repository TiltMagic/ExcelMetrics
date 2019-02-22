import setup
import json
import error_logger
from avg import Report

'''
# dress the excel files
# get summary data
# build summary file
'''

SHEET_TITLE = setup.sheet_title
JOB_TYPES = '"{}"'.format(setup.job_types)


def edit_all_files(sheet_title):
    '''
    Gets files
    Builds file manager objects
    Adds new metric columns and data validation to all column manager objects
    '''
    excel_file_paths = setup.get_file_paths(".xlsx")
    column_managers = setup.get_column_managers(excel_file_paths, sheet_title)
    setup.dress_excel_file(JOB_TYPES, column_managers)


def main():
    '''
    This runs the "whole-shebang"
    '''
    edit_all_files(SHEET_TITLE)
    error_logger.log_it()

    '''
    Comments below are experimenting w/ average file/ data structure
    '''

    # with open("config.json", "r") as data:
    #     things = json.load(data)
    #
    # print(things)

    excel_file_paths = setup.get_file_paths(".xlsx")
    column_managers = setup.get_column_managers(excel_file_paths, SHEET_TITLE)
    manager = column_managers[0]

    report = Report(manager, "Labor $/Hr", "Labor Hours/Unit Area")
    print("\n")
    print(report.all_metrics)

    # data = avg.get_data_structure(manager, "Labor $/Hr")
    # print("\n")
    # # print(data['Residential'])
    # for value in data['Residential']:
    #     print(value + value[0])


if __name__ == "__main__":
    main()
