import timestamp
import logging
import format_file
import os

log_info = []


def logger(statement=None, traceback_error=None):
    '''
    Prints, then appends any errors to log_info list
    '''
    if statement:
        log_info.append(statement)
        print(statement)

    if traceback_error:
        log_info.append(traceback_error)
        if format_file.with_raise:
            print(traceback_error)


def log_the_info():
    '''
    Logs all the info in the log_info list
    '''
    for info in log_info:
        logging.debug(info)

    logging.debug(timestamp.timestamp)
    logging.exception(
        '\n****TRACEBACK ERRORS BELOW (if enabled in config.json - WITH_RAISE = "True")****')


def build_error_dir():
    '''
    Builds log_files dir if it does not already exists
    '''
    cwd = os.getcwd()
    try:
        os.mkdir('{}\log_files'.format(cwd))
    except OSError:
        None
    except Exception as traceback_error:
        statement = "Trouble building 'log_files' directory"
        logger(statement, traceback_error)


def build_log_file():
    '''
    Creates new file in log_files dir and logs all errors to new log file
    '''
    cwd = os.getcwd()
    build_error_dir()
    file_name = '{}\log_files\enterprise-{}.log'.format(cwd, timestamp.timestamp)
    logging.basicConfig(filename=file_name, level=logging.DEBUG)
    log_the_info()
