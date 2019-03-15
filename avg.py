import format_file
import columnmanager
import report
import statistics
import json

with open("config.json", "r") as config_file:
    configs = json.load(config_file)

NEW_COLUMNS = configs['NEW_COLUMNS']
SHEET_TITLE = configs['SHEET_TITLE']
METRIC_TITLES = [column['title'] for column in configs['NEW_COLUMNS']]
JOB_TYPES = configs['JOB_TYPES']
SYSTEMS = configs['SYSTEMS']


final_data = {}


def build_data_set_template(metric_titles=METRIC_TITLES):
    final_data = {}

    for job_type in JOB_TYPES:
        final_data[job_type] = {}
        for metric in METRIC_TITLES:
            final_data[job_type][metric] = {}
            for system in SYSTEMS:
                final_data[job_type][metric][system] = []


    return final_data


def get_all_reports(column_managers, metric_titles=METRIC_TITLES):
    reports = [report.Report(manager, metric_titles).get_final_report() for manager in column_managers]
    return reports

def build_final_data_set(column_managers):
    template = build_data_set_template()
    reports = get_all_reports(column_managers)

    for report in reports:
        for job_type in report:
            if job_type == None:
                continue
            for metric in report[job_type]:
                for system in report[job_type][metric]:
                    if report[job_type][metric][system] == 'None':
                        continue
                    template[job_type][metric][system].append(report[job_type][metric][system])


    return template
    #                 # print(report[value][thing][thang], "\n")

def build_avg_data(column_managers):
    data_set = build_final_data_set(column_managers)


    for job_type in data_set:
        if job_type == None:
            continue
        for metric in data_set[job_type]:
            for system in data_set[job_type][metric]:
                if data_set[job_type][metric][system] == None or \
                   not data_set[job_type][metric][system]:
                    continue
                print("{}:{}:{}: {}".format(job_type,metric,system, data_set[job_type][metric][system]))
                data_set[job_type][metric][system] = statistics.mean(data_set[job_type][metric][system])
                # print(data_set[job_type][metric][system])

    print("\n",data_set)





    '''
    Comments below are experimenting w/ average file/ data structures
    '''

    # excel_file_paths = format_file.get_file_paths(".xlsx")
    #
    # column_managers = []
    #
    # for file_path in excel_file_paths:
    #     column_managers.append(columnmanager.ColumnManager(file_path, SHEET_TITLE))
    #
    #
    # manager = column_managers[0]
    #
    # report = Report(manager, "Labor $/Hr", "Mat Labor.", "Total Hrs")
    # print("\n")
    #
    # for value in report.all_metrics:
    #     print(value, "\n")
