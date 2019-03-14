import format_file
import json
import error_logger
import columnmanager
import avg
from report import Report

'''
# dress the excel files
# get summary data
# build summary file
'''

SHEET_TITLE = format_file.sheet_title
JOB_TYPES = '"{}"'.format(format_file.job_types)




def run_program(sheet_title):
    '''
    Gets files
    Builds file manager objects
    Adds new metric columns and data validation to all column manager objects
    '''
    excel_file_paths = format_file.get_file_paths(".xlsx")

    try:
        format_file.format_excel_files(excel_file_paths, JOB_TYPES)
    except Exception as traceback_error:
        error_logger.logger(traceback_error=traceback_error)

    column_managers = [columnmanager.ColumnManager(path, SHEET_TITLE) for path in excel_file_paths]
    result = avg.build_final_data_set(column_managers)
    print(result)

def main():
    '''
    This runs the entire program
    '''
    run_program(SHEET_TITLE)
    error_logger.build_log_file()


if __name__ == "__main__":
    main()



    '''
    Comments below are experimenting w/ average file/ data structure
    '''
