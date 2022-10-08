from __future__ import annotations

from collections    import defaultdict
from math           import ceil
from openpyxl       import load_workbook
from openpyxl.utils import get_column_letter
from os             import walk
from os.path        import exists
from src.cls        import Capacity, Category, GradeLevel, Group, ParallelSession, Partition, Section, Shift, Student, Type
from typing         import Any, Iterable
from uuid           import UUID, uuid4
from xlsxwriter     import Workbook

import random

MATH = 'Math'
CORE = 'Core science'
ELEC = 'STE elective'
RESEARCH = 'Research'


class WriteAgent:
    DEFAULT_SHEETNAME = 'Sheet1'

    def as_text(value):
        if value is None:
            return value
        if isinstance(value, float) or isinstance(value, int) or (isinstance(value, str) and value.isdigit()):
            return int(value)
        return str(value)
    
    def generate_xlsx(path: str):
        workbook = Workbook(path)
        workbook.close()

    def delete_default_sheet(path: str):
        workbook = load_workbook(path)
        if WriteAgent.DEFAULT_SHEETNAME in workbook.sheetnames:
            workbook.remove(workbook[WriteAgent.DEFAULT_SHEETNAME])
        workbook.save(path)
        workbook.close()

    def realign_columns(path: str, sheet: str):
        workbook = load_workbook(path)
        worksheet = workbook[sheet]
        for column_cells in worksheet.columns:
            worksheet.column_dimensions[
                get_column_letter(column_cells[0].column)
            ].width = max(len(repr(cell.value)) for cell in column_cells)
        workbook.save(path)
        workbook.close()
    
    def to_xlsx(path: str, sheet: str, data: list[list]):
        if not exists(path):
            WriteAgent.generate_xlsx(path)

        workbook = load_workbook(path)
        worksheet = workbook.create_sheet(sheet)
        for row in data:
            worksheet.append(WriteAgent.as_text(cell) for cell in row)
        workbook.save(path)
        workbook.close()

        WriteAgent.delete_default_sheet(path)
        WriteAgent.realign_columns(path, sheet)

    def find_filepath(directory: str, name_template: str) -> str:
        path_template = f'{directory}/{name_template}'
        for index in range(len(next(walk(directory))[2]) + 1):
            if not exists(path_template.format(index)):
                return path_template.format(index)

    def get_sheetnames(path: str):
        workbook = load_workbook(path)
        sheetnames = set(workbook.sheetnames)
        workbook.close()
        return sheetnames
    
    def is_sheet_present(path: str, sheet: str):
        return sheet in WriteAgent.get_sheetnames(path)
    
    def find_sheetname(path: str, name_template: str):
        index = 0
        while WriteAgent.is_sheet_present(path, name_template.format(index)):
            index += 1
        return name_template.format(index)
    
    def read_xlsx(path: str, sheet: str):
        workbook = load_workbook(path, data_only=True)
        worksheet = workbook[sheet]
        data = list(list(WriteAgent.as_text(cell.value) for cell in row) for row in worksheet.rows)
        workbook.close()
        return data


class EncodeAgent:
    shifts: set[Shift]
    types: dict[str, Type]
    groups: dict[str, Group]
    categories: dict[str, Category]
    grade_levels: dict[str, GradeLevel]
    students: defaultdict[GradeLevel, list[Student]]

    YES = 'Y'
    GROUP = 'Group'
    DELIMITER = '||'
    SHIFTS = 'Shifts'
    GRADE_LEVEL = 'Grade level'
    PREVIOUS_YEAR = 'Previous year'
    PREREQUISITES = 'Prerequisites'
    NOT_ALONGSIDE = 'Not alongside'
    CLASSIFICATION = 'Classification'
    INFORMATION = 'General information'

    def __init__(self):
        self.shifts = set()
        self.types = dict()
        self.groups = dict()
        self.categories = dict()
        self.grade_levels = dict()
        self.students = defaultdict(list)

    def __repr__(self):
        return 'Encoding Agent'

    def get_subjects(self, path: str):
        # Encode shifts
        data = WriteAgent.read_xlsx(path, EncodeAgent.SHIFTS)
        for r in range(1, len(data)):
            shift = Shift()
            for c in range(1, 1 + data[r][0]):
                shift.add(Partition(data[r][c]))
            self.shifts.add(shift)

        # Encode grade levels
        data = WriteAgent.read_xlsx(path, EncodeAgent.GRADE_LEVEL)
        for r in range(len(data)):
            grade_level = GradeLevel(data[r][0])
            self.grade_levels[str(grade_level)] = grade_level

        # Encode subject information
        ## Encode names and capacities
        data = WriteAgent.read_xlsx(path, EncodeAgent.INFORMATION)
        for r in range(1, len(data)):
            category = Category(data[r][0], data[r][1] if isinstance(data[r][1], int) else None)
            category.capacity_section = Capacity(data[r][2], data[r][4], data[r][3])
            category.capacity_sections = Capacity(0, data[r][5])
            self.categories[str(category)] = category

        ## Encode linked subjects
        for r in range(1, len(data)):
            key = data[r][6]
            if key in self.categories:
                self.categories[str(Category(data[r][0], data[r][1] if isinstance(data[r][1], int) else None))].link = self.categories[key]

        ## Encode classification
        data = WriteAgent.read_xlsx(path, EncodeAgent.CLASSIFICATION)
        for r in range(1, len(data)):
            for c in range(1, len(data[r]), 3):
                if any(cell == EncodeAgent.YES for cell in data[r][c:c + 3]):
                    type  = data[0][c + 2]
                    if type not in self.types:
                        self.types[type] = Type(type, len(self.types))
                    self.grade_levels[data[0][c]].add(self.types[type], data[0][c + 1] == EncodeAgent.YES, self.categories[data[r][0]])

        ## Encode prerequisites
        data = WriteAgent.read_xlsx(path, EncodeAgent.PREREQUISITES)
        data = WriteAgent.read_xlsx(path, self.PREREQUISITES)
        for r in range(1, len(data)):
            for c in range(2, 2 + data[r][1]):
                prerequisites = tuple(self.categories[s] for s in data[r][c].split(self.DELIMITER))
                self.categories[data[r][0]].add_prerequisites(prerequisites)
        
        ## Encode not alongsides
        data = WriteAgent.read_xlsx(path, self.NOT_ALONGSIDE)
        for r in range(1, len(data)):
            for c in range(2, 2 + data[r][1]):
                self.categories[data[r][0]].add_not_alongside(self.categories[data[r][c]])

    def get_students(self, path: str):
        # Encode groups
        for sheet in WriteAgent.get_sheetnames(path):
            if sheet.find(EncodeAgent.GROUP) != -1:
                data = WriteAgent.read_xlsx(path, sheet)
                for r in range(2, len(data)):
                    group = Group(data[r][0], self.categories[data[0][1]])
                    self.groups[str(group)] = group
        
        # Encode students
        for sheet in WriteAgent.get_sheetnames(path):
            if sheet.find(self.GRADE_LEVEL) != -1:
                data = WriteAgent.read_xlsx(path, sheet)
                for r in range(2, len(data)):
                    grade_level = self.grade_levels[data[r][0]]
                    student = Student(data[r][1], grade_level)

                    c = 2
                    if data[0][c].upper().find(EncodeAgent.GROUP.upper()) != -1:
                        ## Encode group
                        parent = self.categories[data[1][2]]
                        key = f'{parent} {data[r][c]}'
                        if key in self.groups:
                            self.groups[key].add(student)
                            student.group = self.groups[key]
                        c += 1
                    while data[0][c].upper() == self.PREVIOUS_YEAR.upper():
                        student.taken.add(self.categories[data[r][c]])
                        c += 1
                    for d in range(c, len(data[r])):
                        if data[r][d]:
                            student.rankings.add(self.types[data[1][d]], self.categories[data[r][d]])
                    self.students[grade_level].append(student)

    def clean(self):
        for students in self.students.values():
            for student in filter(lambda x: x.rankings.current(self.types[CORE]) in x.rankings.final.ordered[self.types[ELEC]], students):
                final_ranks = student.rankings.final
                start_ranks = student.rankings.start
                category = start_ranks.ordered[self.types[ELEC]].index(student.rankings.initial(self.types[CORE]))
                final_ranks.pop(self.types[ELEC], 'Already the first choice Core elective', category)
                start_ranks.pop(self.types[ELEC], 'Already the first choice Core elective', category)
                
                
class AnalyzeAgent:
    def score(encode_agent: EncodeAgent, show: bool = None):
        total = defaultdict[Type, int](int)
        attained = defaultdict[Type, int](int)
        for students in encode_agent.students.values():
            for student in students:
                for type, _ in filter(lambda x: x[1], student.grade_level.categories):
                    total[type] += 1
                    if type in student.takes and student.rankings.initial(type) == student.takes[type]:
                        attained[type] += 1
        
        scores = dict((type, attained[type] / total[type] * 100) for type in total)
        if show:
            for type in scores:
                print(f'{scores[type]}% [{type}]')
        return scores