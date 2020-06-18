from openpyxl import *


class Spreadsheet:
    def __int__(self, fileName):
        self.fileName = fileName
        self.workbook = load_workbook(filename=self.fileName)

        self.setWorksheets()

        self.clipboard = []

    def setWorksheets(self):
        worksheets = self.workbook.worksheets
        start_found_index = worksheets.index("500 (All)")
        start_core_index = worksheets["510 (All)"]

        foundation_courses_names = worksheets[start_found_index:start_core_index]

        core_courses_names = worksheets[start_core_index:]

        special_topics_names = []

        for course in core_courses_names:
            if course[2] != '0':
                special_topics_names.append(course)
                core_courses_names.remove(course)

        # Create dictionaries for each course with worksheet
        self.foundation_courses = {foundation_courses_names[i]: self.Worksheet(self, foundation_courses_names[i]) for i
                                   in range(0, len(foundation_courses_names))}

        self.core_courses = {core_courses_names[i]: self.Worksheet(self, core_courses_names[i]) for i in
                             range(0, len(core_courses_names))}

        self.special_topics_courses = {special_topics_names[i]: self.Worksheet(self, special_topics_names[i]) for i in
                                       range(0, len(special_topics_names))}

    class Worksheet:
        def __init__(self, current_spreadsheet, current_workbook_name):
            self.current_spreadsheet = current_spreadsheet
            self.current_worksheet = current_spreadsheet.workbook[current_workbook_name]
