import columnmanager
import os
import error_logger
import json
import sys

# read from config file here - config.json
# maybe this should be a seperate file?
# maybe use try block here

with open("config.json", "r") as config_file:
    configs = json.load(config_file)

try:
    sheet_title = configs['SHEET_TITLE']
    job_types = ", ".join(configs['JOB_TYPES'])
    column_color = configs['COLUMN_COLOR']
    with_formulas = configs['WITH_FORMULAS']
    with_raise = configs['WITH_RAISE']
    new_columns = configs['NEW_COLUMNS']
except Exception as traceback_error:
    statement = "Trouble loading from config file"
    error_logger.logger(statement, traceback_error)

# read from config.json ^^^^^^^^^^^^^^^^^^^


def get_file_paths(file_extention):
    '''
    Returns file paths for Excel files only- in current working directory
    '''
    cwd = os.getcwd()
    files = os.listdir(cwd)
    excel_files = [file for file in files if file[len(file)-5:] == file_extention]
    excel_file_paths = [cwd + "\{}".format(file_name) for file_name in excel_files]
    return excel_file_paths


def build_metrics_columns(manager, new_columns=new_columns, row=1):
    """
    Uses excel manager object to build new metrics columns
    """
    try:
        for new_column in new_columns:
            manager.gen_new_metric_column(new_column['title'], new_column['numerator'], new_column['denominator'], with_formulas=with_formulas)
            manager.color_column(new_column['title'], column_color)
    except Exception as traceback_error:
        print(traceback_error)
        statement = "Trouble building metrics column"
        error_logger.logger(statement, traceback_error)


def get_formatted_filepath_status(file_path):
    '''
    Formats filepath into file status output marquee
    '''
    file_name = list(file_path.split("\\"))[-1]
    statement = " Status below for file: {}\n".format(file_name.upper())
    line = "-" * len(statement)
    output = ("{}\n"
              "           {}"
              "           {}".format(line, statement, line))
    return output

def all_new_columns_exist(column_manager):
    '''
    Checks if all new columns specified in json file exist in column_manager's Excel file
    '''
    for new_column in new_columns:
        if column_manager.value_exists(new_column['title']):
            pass
        else:
            return False

    return True


def format_excel_file(JOB_TYPES, file_path):
    '''
    Format excel file to final result given file path
    '''
    try:
        column_manager = columnmanager.ColumnManager(file_path, sheet_title)
    except Exception as traceback_error:
        statement_1 = "Problem manipulating file from file path"
        statement_2 = "Make sure Excel file is formatted properly- maybe try to re-export Extension from Enterprise\n"
        error_logger.logger("{}\n{}".format(statement_1, statement_2), traceback_error)


    if column_manager.group_by_exists():
        try:
            build_metrics_columns(column_manager)
            column_manager.add_validation(JOB_TYPES)
            column_manager.save_doc()
        except Exception as traceback_error:
            # location = manager.get_error_location()
            statement = "Problem building metrics columns for {}".format(location)
            error_logger.logger(statement, traceback_error)
    elif all_new_columns_exist(column_manager) and column_manager.sheet['A2'].value == 'Group By':
        formatted_check_symbol = "  /\n\/"
        error_logger.logger(formatted_check_symbol)
        # error_logger.logger('Good')
    else:
        statement_1 = "Excel file was NOT formatted"
        statement_2 = ("Make sure 'Group By' text is located in cell A1- and rest of Excel file is formatted properly"
                      "\n..OR re-export Extension from Enterprise and try again")
        formatted_fail_symbol = "\  /\n \/\n /\\\n/  \\"
        error_logger.logger(statement_1)
        error_logger.logger(statement_2)
        error_logger.logger(formatted_fail_symbol)


def format_excel_files(file_paths, JOB_TYPES):
    '''
    Format all excel files to final result given list of file paths
    '''
    # try:
    for file_path in file_paths:
        status_output = get_formatted_filepath_status(file_path)
        error_logger.logger(status_output)
        try:
            format_excel_file(JOB_TYPES, file_path)

        except Exception as traceback_error:
            statement_1 = "Problem manipulating file from filepath"
            statement_2 = "Make sure Excel file is formatted properly- maybe try to re-export from Enterprise\n"
            error_logger.logger("{}\n{}".format(statement_1, statement_2), traceback_error)
