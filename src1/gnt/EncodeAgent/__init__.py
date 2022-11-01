# There imports are needed to have the EncodeAgent run
from src import *

# Encode agent implementation in Python 3.10.6
class EncodeAgent:
  shifts: set[Shift]
  courses: dict[str, Course]
  course_types: dict[str, CourseType]
  grade_levels: dict[str, GradeLevel]
  research_groups: dict[str, ResearchGroup]
  students: defaultdict[GradeLevel, list[Student]]
  
  # Initialize the instance of this class
  def __init__(self):
    self.shifts = set()
    self.courses = dict()
    self.course_types = dict()
    self.grade_levels = dict()
    self.research_groups = dict()
    self.students = defaultdict(list)
    
  # Given a .xlsx file, encode the courses in the workbook
  def encode_courses(self, path: str):
    # Encode shifts
    sheet = read_xlsx(path, 'Shifts')
    for r in range(1, len(sheet)):
      shift = Shift()
      for c in range(1, 1 + sheet[r][0]):
        shift.add(Session(sheet[r][c]))
      self.shifts.add(shift)
      
    # Encode grade levels
    sheet = read_xlsx(path, 'Grade levels')
    for r in range(len(sheet)):
      grade_level = GradeLevel(sheet[r][0])
      self.grade_levels[str(grade_level)] = grade_level
      
    # Encode course names and capacities
    sheet = read_xlsx(path, 'Course names')
    for r in range(1, len(sheet)):
      course = Course(sheet[r][0], sheet[r][1])
      self.courses[str(course)] = course
      
    sheet = read_xlsx(path, 'Course capacities')
    for r in range(1, len(sheet)):
      course = self.courses[sheet[r][0]]
      course.capacity_section = Capacity(
        sheet[r][1], sheet[r][2], sheet[r][3])
      course.capacity_sections = Capacity(0, 0, sheet[r][4])
      
    # Encode linked courses
    sheet = read_xlsx(path, 'Course links')
    for r in range(1, len(sheet)):
      self.courses[sheet[r][0]].linked_to = self.courses[sheet[r][1]]
        
    # Encode course types
    
    sheet = read_xlsx(path, 'Course classification')
    for r in range(1, len(sheet)):
      for c in range(1, len(sheet[r]), 3):
        if any(cell == 'Y' for cell in sheet[r][c:c + 3]):
          course_type = sheet[0][c + 2]
          if course_type not in self.course_types:
            self.course_types[course_type] = CourseType(
              course_type, len(self.course_types))
          self.grade_levels[sheet[0][c]].add(
            self.course_types[course_type],
            sheet[0][c + 1] == 'Y',
            self.courses[sheet[r][0]])
    
    # Encode prerequisite courses
    sheet = read_xlsx(path, 'Course prerequisites')
    for r in range(1, len(sheet)):
      for c in range(2, 2 + sheet[r][1]):
        self.courses[sheet[r][0]].prerequisites.append(set(
          self.courses[course] for course in sheet[r][c].split('||')))
        
    # Encode courses that cannot be taken alongside a course
    sheet = read_xlsx(path, 'Course not alongside')
    for r in range(1, len(sheet)):
      for c in range(2, 2 + sheet[r][1]):
        self.courses[sheet[r][0]].not_alongside.add(
          self.courses[sheet[r][c]])
    
  # Given a .xlsx file, encode the students in the workbook
  def encode_students(self, path: str):
    # Encode research groups
    sheet = read_xlsx(path, 'Research groups')
    for r in range(2, len(sheet)):
      research_group = ResearchGroup(
        self.courses[sheet[0][1]], sheet[r][0])
      self.research_groups[str(research_group)] = research_group

    # Encode students
    for grade_level_alias in self.grade_levels:
      sheet = read_xlsx(path, grade_level_alias)
      for r in range(2, len(sheet)):
        grade_level = self.grade_levels[sheet[r][0]]
        student = Student(sheet[r][1], grade_level)
        
        c = 2
        if sheet[0][c] == 'Groups':
          research_course = self.courses[sheet[1][2]]
          key = f'{research_course} {sheet[r][c]}'
          if key in self.research_groups:
            self.research_groups[key].add(student)
            student.research_group = self.research_groups[key]
          c += 1
        while sheet[0][c] == 'Previous year':
          student.taken.add(self.courses[sheet[r][c]])
          c += 1
        for d in range(c, len(sheet[r])):
          if sheet[r][d]:
            student.rankings.add(
              self.course_types[sheet[1][d]], 
              self.courses[sheet[r][d]])
        self.students[grade_level].append(student)
      self.students[self.grade_levels[grade_level_alias]].sort()