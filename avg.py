import format_file
import columnmanager
import report
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

    for job in JOB_TYPES:
        final_data[job] = {}
        for metric in METRIC_TITLES:
            final_data[job][metric] = {}
            for system in SYSTEMS:
                final_data[job][metric][system] = []


    return final_data


def get_all_reports(column_managers, metric_titles=METRIC_TITLES):
    reports = [report.Report(manager, metric_titles).get_final_report() for manager in column_managers]
    return reports

def build_final_data_set(column_managers):
    template = build_data_set_template()
    reports = get_all_reports(column_managers)

    for report in reports:
        print(report)
        # for job_type in report:
        #     for metric in job_type:
        #         for system in metric:
                    # template[job_type][metric][system].append(report[job_type][metric][system])
                    # print(template[job_type][metric])

    return template


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
