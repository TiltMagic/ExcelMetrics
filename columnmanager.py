from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
import openpyxl
import os
import timestamp
import error_logger
import json

with open("config.json", "r") as config_file:
    configs = json.load(config_file)

# Change these variables so they're not so verbose
labor_dollar_hr_title = configs['LABOR $/Hr Title']
labor_dollar_unit_area_title = configs['LABOR HOURS/UNIT AREA']


class ColumnManager:

    def __init__(self, file_location, sheet_name):
        self.workbook = self.load_workbook(file_location)
        self.sheet = self.get_sheet(self.workbook, sheet_name)
        self.file_location = file_location
        self.file_name = list(file_location.split("\\"))[-1]

    def load_workbook(self, file_location):
        """
        Loads workbook object from excel file
        """
        try:
            return openpyxl.load_workbook(file_location)
        except Exception as traceback_error:
            statement = "Problem laoding workbook from {}".format(file_location)
            error_logger.logger(statement, traceback_error)

    def save_doc(self):
        # Change to enable saving to any location *args
        """
        Saves ColumnManager changes to excel workbook
        """
        self.workbook.save(self.file_location)

    def get_sheet(self, workbook_name, sheet_name):
        """
        Retrieves sheet object from workbook object
        """
        try:
            return workbook_name[sheet_name]
        except Exception as traceback_error:
            statement = "Problem finding {} worksheet from {} workbook".format(
                sheet_name, workbook_name)
            error_logger.logger(statement, traceback_error)

    def make_new_column(self, title):
        """
        Create column w/ given title
        """
        column_titles = self.get_column_titles()

        if title in column_titles:
            statement = "Column title {} already exists for file {}".format(title, self.file_name)
            error_logger.logger(statement)
        else:
            try:
                max_column = self.sheet.max_column
                self.sheet.cell(row=1, column=max_column+1).value = title
            except Exception as traceback_error:
                statement = "Something went wrong with building column {}".format(title)
                self.logger(statement, traceback_error)

    def get_column_titles(self, row=1):
        """
        Returns list with column title
        """

        # Soluttion to add metrics columns if conditional formatting row exists
        # would probably be here
        max_column = self.sheet.max_column

        try:
            column_titles = [self.sheet.cell(row=row, column=column).value
                             for column in range(1, max_column+1)]
            return column_titles
        except Exception as traceback_error:
            statement = "Problem retrieving column titles"
            error_logger.logger(statement, traceback_error)

    def get_column_titles_with_index(self, row=1):
        """
        Returns dict with column title and column index
        """
        max_column = self.sheet.max_column
        try:
            column_titles = {self.sheet.cell(row=row, column=column).value: column
                             for column in range(1, max_column+1)}
            return column_titles
        except Exception as traceback_error:
            statement = "Problem retrieving column titles"
            error_logger.logger(statement, traceback_error)

    def get_column_index(self, title, row=1):
        """
        Return index number for given column title
        """
        try:
            columns = self.get_column_titles_with_index(row)
            return columns[title]
        except Exception as traceback_error:
            statement = "Problem finding column titled {}".format(title)
            error_logger.logger(statement, traceback_error)

    def get_column_cells(self, title, row=1):
        """
        Returns list of cell objects from given column title
        """
        index = self.get_column_index(title, row)
        try:
            return [cell for cell in list(self.sheet.columns)[index - 1]]
        except Exception as traceback_error:
            statement = "Problem finding cell for column w/ title {}".format(title)
            error_logger.logger(statement, traceback_error)

    def get_column_cell_names(self, title, row=1):
        """
        Return a list of cell names (ex."A32") for given column
        """
        cells = self.get_column_cells(title, row)
        try:
            cell_names = ["{}{}".format(cell.column, cell.row)
                          for cell in cells]
            return cell_names
        except Execption as traceback_error:
            statement = "Problem building cell name from {} column".format(title)
            error_logger.logger(statement, traceback_error)

    def get_column_values(self, title, row=1):
        """
        Returns column values given title
        """
        try:
            cells = self.get_column_cells(title, row)
            values = [cell.value for cell in cells]
            return values
        except Exception as traceback_error:
            statement = "Problem retrieving cells for column {}".format(title)
            error_logger.logger(statement, traceback_error)

    def set_column_values(self, title, values, row=2):
        """
        Sets cell values for given column
        """
        try:
            index = self.get_column_index(title)
            cells_to_set = self.get_column_cells(title)[1:]
        except Exception as traceback_error:
            statement = "Problem with inputs to set_column_values"
            error_logger.logger(statement, traceback_error)

        for i in range(len(cells_to_set)):
            try:
                self.sheet.cell(row=row, column=index).value = values[i]
            except Exception as traceback_error:
                statement = "Problem setting cell values"
                error_logger.logger(statement, traceback_error)
            row += 1

    def print_column_values(self, title, row=1):
        """
        Prints out cell values
        """
        cells = self.get_column_cells(title, row)
        for cell in cells:
            print(cell.value)

    def divide_columns_values(self, first_title, second_title, row=1):
        # maybe change return type to tuple
        """
        Divides two columns and return list w/ resulting values
        """
        numerator = self.get_column_values(first_title, row=row)
        denominator = self.get_column_values(second_title, row=row)
        result = []
        for i in range(1, len(numerator)):
            try:
                result.append(numerator[i]/denominator[i])
            except (TypeError, ZeroDivisionError) as e:
                # Should the above exception be logged somehow?
                result.append("None")

        return result

    def divide_columns_formula(self, first_title, second_title, row=1):
        """
        Return list of divided cell values from given columns
        """
        numerator_names = self.get_column_cell_names(first_title, row=row)
        denominator_names = self.get_column_cell_names(second_title, row=row)
        divided_formulas = ["={}/{}".format(numerator_names[i],
                                            denominator_names[i])
                            for i
                            in range(1, len(numerator_names))]
        return divided_formulas

    def gen_labordollar_perhour_column(self, with_formulas=False, clm_title=labor_dollar_hr_title, row=1):
        """
        Generates Labor $/Hour column with values
        with_formulas parameter toggles excel formulas shown in metrics column
        Edit with_fromulas in config file.
        """
        self.make_new_column(clm_title)

        try:
            if with_formulas == True:
                new_values = self.divide_columns_formula("Total Labor $",
                                                         "Total Hrs", row=row)
            elif with_formulas == False:
                new_values = self.divide_columns_values("Total Labor $", "Total Hrs", row=row)

            self.set_column_values(clm_title, new_values, row=row+1)
        except Exception as traceback_error:
            statement = "Trouble with generating Labor $ / Hour column"
            error_logger.logger(statement, traceback_error)

    def gen_laborhours_unitarea(self, with_formulas=False, clm_title=labor_dollar_unit_area_title, row=1):
        """
        Generates Labor Hours/Unit Area column w/ values
        with_formulas parameter toggles excel formulas shown in metrics column
        Edit with_formulas in config file.
        """
        # Should there be a try block here?
        self.make_new_column(clm_title)
        if with_formulas == True:
            new_values = self.divide_columns_formula("Total Hrs", "Unit", row=row)
        elif with_formulas == False:
            new_values = self.divide_columns_values("Total Hrs", "Unit", row=row)

        self.set_column_values(clm_title, new_values, row=row+1)

    def color_column(self, title, color="B3FFB3"):
        """
        read color from config file, keep B3FFB3 as default
        """
        cells = self.get_column_cells(title)
        for cell in cells:
            cell.fill = PatternFill(start_color=color,
                                    end_color=color,
                                    fill_type="solid")

    def add_validation(self, validation_string):
        '''
        Inserts a new row then adds a data validation drop down list to cell "A1".
        validation_string parameter contains list of values added to drop down list.
        '''
        try:
            self.sheet.insert_rows(1)
            dv = DataValidation(type="list", formula1=validation_string, allow_blank=True)
            self.sheet.add_data_validation(dv)
            cell = self.sheet["A1"]
            dv.add(cell)
        except Exception as traceback_error:
            statement = "Trouble with adding validation drop-down menu"
            error_logger.logger(statement, traceback_error)

    def get_jobtype(self):
        return self.sheet.cell(row=1, column=1).value

    def get_error_location(self):
        error_statement = " Errors above found with file: {}\n".format(self.file_name.upper())
        line = "-" * len(error_statement)
        output = ("{}\n"
                  "           {}"
                  "           {}".format(line, error_statement, line))
        # error_logger.logger(output)
        return output
