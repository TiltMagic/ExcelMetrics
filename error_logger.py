import timestamp
import logging
import os

log_info = []


def logger(statement):
    '''
    Prints, then appends any errors to log_info list
    '''
    print(statement)
    log_info.append(statement)


def log_the_info():
    '''
    Logs all the info in the log_info list
    '''
    for info in log_info:
        logging.debug(info)

    logging.debug(timestamp.timestamp)
    logging.exception(
        '****TRACEBACK ERRORS BELOW (if enabled in config.json - WITH_RAISE = "True")****')


def build_error_dir():
    '''
    Builds log_files dir if it does not already exists
    '''
    cwd = os.getcwd()
    try:
        os.mkdir('{}\log_files'.format(cwd))
    except:
        # maybe add something else here
        None


def log_it():
    '''
    Creates new file in log_files dir and logs all errors to new log file
    '''
    cwd = os.getcwd()
    build_error_dir()
    file_name = '{}\log_files\enterprise-{}.log'.format(cwd, timestamp.timestamp)
    logging.basicConfig(filename=file_name, level=logging.DEBUG)
    log_the_info()
