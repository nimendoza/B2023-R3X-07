# There imports are needed to have the EncodeAgent run
from src.cls import *
from src.gnt.ReadWriteAgent import ReadWriteAgent

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
    data = ReadWriteAgent.read_xlsx(path, 'Shifts')
    for r in range(1, len(data)):
      shift = Shift()
      for c in range(1, 1 + data[r][0]):
        shift.add(Session(data[r][c]))
      self.shifts.add(shift)
      
    # Encode grade levels
    data = ReadWriteAgent.read_xlsx(path, 'Grade level')
    for r in range(len(data)):
      grade_level = GradeLevel(data[r][0])
      self.grade_levels[str(grade_level)] = grade_level
      
    # Encode course names and capacities
    data = ReadWriteAgent.read_xlsx(path, 'General information')
    for r in range(1, len(data)):
      alias = data[r][0]
      difficulty_level = data[r][1] if isinstance(data[r][1], int) else 0
      course = Course(alias, difficulty_level)
      
      minimum = data[r][2]
      maximum = data[r][4]
      course.capacity_section = Capacity(minimum, maximum)
      
      maximum = data[r][5]
      course.capacity_sections = Capacity(0, maximum)
      self.courses[str(course)] = course
      
    # Encode linked courses
    for r in range(1, len(data)):
      course = data[r][6]
      if course in self.courses:
        alias = data[r][0]
        difficulty_level = data[r][1] if isinstance(data[r][1], int) else 0
        linked_course = Course(alias, difficulty_level)
        self.courses[str(linked_course)].linked_to = self.courses[course]
        
    # Encode course types
    data = ReadWriteAgent.read_xlsx(path, 'Classification')
    for r in range(1, len(data)):
      for c in range(1, len(data[r]), 3):
        if any(cell == 'Y' for cell in data[r][c:c + 3]):
          alias = data[0][c + 2]
          if alias not in self.course_types:
            self.course_types[alias] = CourseType(
              alias, 
              len(self.course_types)
            )
          self.grade_levels[data[0][c]].add(
            self.course_types[alias],
            data[0][c + 1] == 'Y',
            self.courses[data[r][0]]
          )
    
    # Encode prerequisite courses
    data = ReadWriteAgent.read_xlsx(path, 'Prerequisites')
    for r in range(1, len(data)):
      for c in range(2, 2 + data[r][1]):
        prerequisites = tuple(
          self.courses[s] 
            for s in data[r][c].split('||')
        )
        self.courses[data[r][0]].add_prerequisites(prerequisites)
        
    # Encode courses that cannot be taken alongside a course
    data = ReadWriteAgent.read_xlsx(path, 'Not alongside')
    for r in range(1, len(data)):
      for c in range(2, 2 + data[r][1]):
        self.courses[data[r][0]].add_not_alongside(
          self.courses[data[r][c]]
        )
    
  # Given a .xlsx file, encode the students in the workbook
  def encode_students(self, path: str):
    # Encode research groups
    for sheet in filter(
      lambda sheet: sheet.find('Group') != -1,
      ReadWriteAgent.get_sheetnames(path)
    ):
      data = ReadWriteAgent.read_xlsx(path, sheet)
      for r in range(2, len(data)):
        research_group = ResearchGroup(
          data[r][0], 
          self.courses[data[0][1]]
        )
        self.research_groups[str(research_group)] = research_group

    # Encode students
    for sheet in filter(
      lambda sheet: sheet.find('Grade level') != -1,
      ReadWriteAgent.get_sheetnames(path)
    ):
      data = ReadWriteAgent.read_xlsx(path, sheet)
      for r in range(2, len(data)):
        grade_level = self.grade_levels[data[r][0]]
        student = Student(data[r][1], grade_level)
        
        c = 2
        
        # Encode research group if it appears in the column header
        if data[0][c].find('Group') != -1:
          research_course = self.courses[data[1][2]]
          key = f'{research_course} {data[r][c]}'
          if key in self.research_groups:
            self.research_groups[key].add(student)
            student.research_group = self.research_groups[key]
          c += 1
        
        # Encode, if indicated, the courses taken by tyhe student in
        # their previous year
        while data[0][c] == 'Previous year':
          student.courses_taken.add(self.courses[data[r][c]])
          c += 1
        
        # Encode the student's course rankings
        for d in range(c, len(data[r])):
          if data[r][d]:
            student.rankings.add(
              self.course_types[data[1][d]],
              self.courses[data[r][d]]
            )
        
        # Add the student to the list of students
        self.students[grade_level].append(student)