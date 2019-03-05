import timestamp
import logging
import setup
import os

log_info = []


def logger(statement, traceback_error=""):
    '''
    Prints, then appends any errors to log_info list
    '''
    print(statement)
    log_info.append(statement)
    if setup.with_raise:
        log_info.append(traceback_error)


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
        print(statement)
        print(traceback_error)


def log_it():
    '''
    Creates new file in log_files dir and logs all errors to new log file
    '''
    cwd = os.getcwd()
    build_error_dir()
    file_name = '{}\log_files\enterprise-{}.log'.format(cwd, timestamp.timestamp)
    logging.basicConfig(filename=file_name, level=logging.DEBUG)
    log_the_info()
