
import setup

'''
This module contains functions for colecting metrics data
and calculating averages
'''


class Report:
    def __init__(self, column_manager, *column_titles):
        self.column_manager = column_manager
        self.column_titles = column_titles
        self.get_job_type()
        self.get_all_metrics()

    def get_job_type(self):
        self.job_type = self.column_manager.get_jobtype()

    def get_data_structure(self, column_title):
        """
        Build data structure with systems and corresponding values
        """
        file_data = {column_title: {}}
        systems = setup.get_base_bid_values(self.column_manager, "System")
        values = setup.get_base_bid_values(self.column_manager, column_title)

        if len(systems) == len(values):
            try:
                system_values = list(zip(systems, values))

                for system_value in system_values:
                    file_data[column_title][system_value[0]] = system_value[1]
            except error as e:
                print(e.message + "&&&&&&&&&&&&&&&&")
                print("Values list and systems list are not the same length")

        return file_data

    def get_all_metrics(self):
        self.all_metrics = [self.get_data_structure(title) for title in self.column_titles]
