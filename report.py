import columnmanager
import format_file
import error_logger


class Report:
    def __init__(self, column_manager, column_titles):
        self.column_manager = column_manager
        self.column_titles = column_titles
        self.job_type = self.get_job_type()
        # self.final_report = self.get_final_report()

    def get_job_type(self):
        return self.column_manager.get_jobtype()


    def get_data_structure(self, column_title):
        """
        Build data structure with systems and corresponding values
        """
        file_data = {column_title: {}}
        # Note use of column manager method call below
        systems = columnmanager.ColumnManager.get_base_bid_values(self.column_manager, "System")
        values = columnmanager.ColumnManager.get_base_bid_values(self.column_manager, column_title)

        if len(systems) == len(values):
            try:
                system_values = list(zip(systems, values))

                for system_value in system_values:
                    file_data[column_title][system_value[0]] = system_value[1]
            except Exception as traceback_error:
                print(traceback_error.message)
                print("Values list and systems list are not the same length")

        return file_data

    def get_final_report(self):
        final_data = {self.job_type: {}}
        for title in self.column_titles:
            data = self.get_data_structure(title)
            final_data[self.job_type][title] = data[title]

        return final_data
