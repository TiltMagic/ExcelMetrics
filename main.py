import format_file
import openpyxl
import json
import error_logger
import columnmanager
import avg
import os
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

    avg.show_avg_marquee()
    column_managers = [columnmanager.ColumnManager(path, SHEET_TITLE) for path in excel_file_paths]

    result = avg.build_avg_data(column_managers)

    avg_filename = 'avg_data'
    cwd = os.getcwd()
    try:
        os.mkdir("{}\\{}".format(cwd, avg_filename))
    except OSError:
        None
    except Exception as traceback_error:
        statement = "Trouble building 'avg_data' directory"
        logger(statement, traceback_error)

    try:
        wb = openpyxl.Workbook()
        dest_filename = '{}\\{}\\{}.xlsx'.format(cwd, avg_filename, avg_filename)

        ws1 = wb.active
        ws1.title = "Averages"

        wb.save(filename=dest_filename)
        print('{} has been generated'.format(avg_filename))
    except Exception as traceback_error:
        error_logger.logger("Trouble building {} file".format(dest_filename), traceback_error)

    avg_manager = columnmanager.ColumnManager(dest_filename, "Averages")
    avg_manager.make_new_column('Job Type', column_start=0)
    avg_manager.make_new_column('Metric')
    avg_manager.make_new_column('System')
    avg_manager.make_new_column('Value')

    try:
        row = 2
        for job_type in result:
            for metric in result[job_type]:
                for system in result[job_type][metric]:
                    avg_manager.sheet['A{}'.format(row)].value = str(job_type)
                    avg_manager.sheet['B{}'.format(row)].value = str(metric)
                    avg_manager.sheet['C{}'.format(row)].value = str(system)
                    if result[job_type][metric][system]:
                        avg_manager.sheet['D{}'.format(
                            row)].value = result[job_type][metric][system]
                    row += 1
    except Exception as traceback_error:
        error_logger.logger('Trouble formatting {} with values'.format(
            dest_filename), traceback_error)

    avg_manager.save_doc()


def main():
    '''
    This runs the entire program
    '''
    run_program(SHEET_TITLE)
    error_logger.build_log_file()


if __name__ == "__main__":
    main()
