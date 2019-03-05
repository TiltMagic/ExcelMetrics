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
    cwd = os.getcwd()
    files = os.listdir(cwd)
    excel_files = [file for file in files if file[len(file)-5:] == file_extention]
    excel_file_paths = [cwd + "\{}".format(file_name) for file_name in excel_files]
    return excel_file_paths


def get_column_managers(excel_file_paths, SHEET_TITLE):
    """
    Return list of manager objects for each excel, given file paths
    """
    try:
        return [columnmanager.ColumnManager(file_path, SHEET_TITLE)
                for file_path in excel_file_paths]
    except Exception as traceback_error:
        statement = "Problem getting column managers"
        error_logger.logger(statement, traceback_error)


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
        # Allow color_column method to take multiple column arguments *args
        manager.color_column("Labor $/Hr", column_color)
        manager.color_column("Labor Hours/Unit Area", column_color)
    except Exception as traceback_error:
        print(traceback_error)
        statement = "Trouble building metrics column"
        error_logger.logger(statement, traceback_error)


def dress_excel_file(JOB_TYPES, column_managers):
    """
    Adds datavalidation and metrics columns to given
    list of excel file manager objects
    """
    for manager in column_managers:
        # this try block prevents traceback errors from being shown in console?
        try:
            build_metrics_columns(manager)
            manager.add_validation(JOB_TYPES)
            manager.save_doc()
            # location = manager.get_error_location()
            # print(location)
        except Exception as traceback_error:
            location = manager.get_error_location()
            print(location)
            statement = "Problem with column manager named {}\n{}".format(manager, location)
            error_logger.logger(statement, traceback_error)

            if with_raise:
                error_logger.log_it()
                raise
