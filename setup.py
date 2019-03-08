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


def get_base_bid_values(columnmanager, title, row=2):
    '''
    Returns list of all systems pertaining to Base Bid only, under "System Column"
    This allows metrics to be calculated for Base Bid values only and not Alternates
    '''

    bid_item_values = columnmanager.get_column_values("Bid Item", row=row)
    system_values = columnmanager.get_column_values(title, row=row)

    results = []

    for value in range(len(bid_item_values)):
        if bid_item_values[value] == " || Base Bid":
            results.append(system_values[value])
    return results


def build_metrics_columns(manager, row=1):
    """
    Uses excel manager object to build new metrics columns
    """
    try:
        manager.gen_labordollar_perhour_column(with_formulas, row=row)
        manager.gen_laborhours_unitarea(with_formulas, row=row)
        manager.color_column("Labor $/Hr", column_color)
        manager.color_column("Labor Hours/Unit Area", column_color)
        error_logger.logger('Metrics columns added')
    except Exception as traceback_error:
        print(traceback_error)
        statement = "Trouble building metrics column"
        error_logger.logger(statement, traceback_error)


def get_formatted_filepath_status(file_path):
    '''
    Formats filepath into file status output
    '''
    file_name = list(file_path.split("\\"))[-1]
    statement = " Status below for file: {}\n".format(file_name.upper())
    line = "-" * len(statement)
    output = ("{}\n"
              "           {}"
              "           {}".format(line, statement, line))
    return output


def format_excel_file(JOB_TYPES, file_path):
    '''
    Fromat excel file to final result given file path
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
    elif column_manager.value_exists('Labor $/Hr') and column_manager.value_exists('Labor Hours/Unit Area') and column_manager.sheet['A2'].value == 'Group By':
        formatted_check_symbol = "  /\n\/"
        error_logger.logger(formatted_check_symbol)
        # error_logger.logger('Good')
    else:
        statement_1 = "Excel file was NOT formatted"
        statement_2 = ("Make sure 'Group By' text is located in cell A1- and rest of Excel file is formatted properly"
                      "\n..OR re-export extension from Enterprise and try again")
        error_logger.logger(statement_1)
        error_logger.logger(statement_2)


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
