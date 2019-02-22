import columnmanager
import os
import datetime

'''
All of this should probably be in error_logger.ph
'''

cwd = os.getcwd()
# files = os.listdir(cwd)


def get_time():
    '''Get current timestamp'''
    time = datetime.datetime.now().strftime("%A_%d_%B_%Y_%I_%M%p_")
    return time


# def master_list():
#     """Build master list of tuples with totals for both average columns"""
#     pass


# def get_averages(filename):
#     """return tuple of file averages"""
#     pass


timestamp = get_time()

# os.system("nul > average-{}.xlsx".format(timestamp))


# name file with timestamp
# create new excel file
