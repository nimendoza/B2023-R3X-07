# These imports are needed to have the abstractions run
from __future__ import annotations
from collections import defaultdict
from typing import Iterable

# Capacity abstraction implemented in Python 3.10.6
class Capacity:
  minimum: int
  maximum: int
  
  # Initialize the instance of this class
  def __init__(self, minimum: int, maximum: int):
    self.minimum = minimum
    self.maximum = maximum
    
# Session abstraction implementation in Python 3.10.6
class Session:
  alias: str
  
  # Initialize the instance of this class
  def __init__(self, alias: str):
    self.alias = alias
    
  # String representation of any instance of this class
  def __repr__(self):
    return self.alias
  
  # Way to distinguish whether an instance of this class is “lesser than”
  # another instance of this class
  def __lt__(self, other: Session):
    return self.alias < other.alias
  
# Shift abstraction implementation in Python 3.10.6
class Shift:
  sessions: set[Session]
  
  # Initialize the instance of this class
  def __init__(self):
    self.sessions = set()
    
  # String representation of any instance of this class
  def __repr__(self):
    return ''.join(map(str, self.sessions))
  
  # Add a session to an instance of this class
  def add(self, session: Session):
    self.sessions.add(session)
    
# Parallel session abstraction implementation in Python 3.10.6
class ParallelSession:
  shift: Shift
  session: Session
  index: int
  
  # Initialize the instance of this class
  def __init__(self, shift: Shift, session: Session, index: int):
    self.shift = shift
    self.session = session
    self.index = index
    
  # String representation of any instance of this class
  def __repr__(self):
    return '{}{}'.format(self.session, self.index or '')
  
# Course type abstraction implementation in Python 3.10.6
class CourseType:
  alias: str
  order: int
  
  # Initialize the instance of this class
  def __init__(self, alias: str, order: int):
    self.alias = alias
    self.order = order
    
  # String representation of any instance of this class
  def __repr__(self):
    return self.alias
  
  # Way to distinguish whether an instance of this class is “lesser than”
  # another instance of this class
  def __lt__(self, other: CourseType):
    return self.order < other.order
  
# Grade level abstraction implementation in Python 3.10.6
class GradeLevel:
  alias: int
  courses: defaultdict[tuple[CourseType, bool], set[Course]]
  
  # Initialize the instance of this class
  def __init__(self, alias: int):
    self.alias = alias
    self.courses = defaultdict(set)
    
  # String representation of any instance of this class
  def __repr__(self):
    return f'Grade {self.alias}'
  

  # Add a course, given its course type and whether it is ranked, to an
  # instance of this class
  def add(self, course_type: CourseType, ranked: bool, course: Course):
    self.courses[(course_type, ranked)].add(course)
    
# Ranking abstraction implementation in Python 3.10.6
class Ranking:
  ordered_courses: dict[CourseType, list[Course]]
  reason_rejected: dict[CourseType, dict[Course, str]]
  
  # Initialize the instance of this class
  def __init__(self, grade_level: GradeLevel):
    self.ordered_courses = dict(
      (course_type, list())
        for course_type, ranked in grade_level.courses
          if ranked
    )
    self.reason_rejected = dict(
      (course_type, dict())
        for course_type, ranked in grade_level.courses
          if ranked
    )
    
  # Given the course type, determine the number of courses ranked
  def len(self, course_type: CourseType):
    return len(self.ordered_courses[course_type])
  
  # Add a course to the rankings of the given course type. A last-in-last-
  # out (LILO) basis is observed in the rankings.
  def add(self, course_type: CourseType, course: Course):
    self.ordered_courses[course_type].append(course)
    
  # Remove a course of a given course type, and note the reason to why
  # it was removed from the rankings
  def pop(self, course_type: CourseType, reason: str, index: int):
    course = self.ordered_courses[course_type].pop(index)
    self.reason_rejected[course_type][course] = reason
    
# Rankings abstraction implementation in Python 3.10.6
class Rankings:
  start: Ranking
  final: Ranking
  owner: Student
  
  # Initialize the instance of this class
  def __init__(self, owner: Student):
    self.owner = owner
    self.start = Ranking(owner.grade_level)
    self.final = Ranking(owner.grade_level)
    

  # Add a course of a given course type to both start and final rankings if
  # and only if the owner is qualified to take the course
  def add(self, course_type: CourseType, course: Course):
    if course.qualified(self.owner):
      self.start.add(course_type, course)
      self.final.add(course_type, course)
      
  # Return the index-th (or, if index is not given, first) course
  # in the initial (“start”) rankings of a given course type. 
  # If the ranking is empty, return None
  def initial(self, course_type: CourseType, index: int | None = None):
    if self.start.len(course_type) == 0:
      return None
    return self.start.ordered_courses[course_type][index or 0]
  
  # Return the index-th (or, if index is not given, first) course in
  # the current (“final”) rankings of a given course type. If the
  # ranking is empty, “reset” the rankings according to whatever
  # courses the student (“owner”) is qualified to take
  def current(self, course_type: CourseType, index: int | None = None):
    if self.final.len(course_type) == 0:
      for course in self.owner.grade_level.courses[(course_type, True)]:
        if course.qualified(self.owner):
          self.final.add(course_type, course)
    return self.final.ordered_courses[course_type][index or 0]
  
# Student abstraction implementation in Python 3.10.6
class Student:
  alias: str
  grade_level: GradeLevel
  rankings: Rankings
  research_group: ResearchGroup | None
  courses_taken: set[Course]
  courses_taking: dict[CourseType, Course]
  attending_sessions: set[Session]
  attending_sections: dict[Course, Section]
  __shift: Shift | None
  
  # Initialize the instance of this class
  def __init__(self, alias: str, grade_level: GradeLevel):
    self.alias = alias
    self.grade_level = grade_level
    self.rankings = Rankings(self)
    self.research_group = None
    self.courses_taken = set()
    self.courses_taking = dict()
    self.attending_sessions = set()
    self.attending_sections = dict()
    self.__shift = None
    

  # String representation of any instance of this class
  def __repr__(self):
    return f'{self.grade_level}-{self.alias}'
  
  # This property was implemented instead of simply renaming “__shift” to
  # “shift” due to additional steps done when assigning a value to “shift”
  @property
  def shift(self):
    return self.__shift
  
  # Every time the “shift” value is changed, also change the affected
  # entities’ (research group, sections attended) shifts in the case
  # that they differ from the assigned value.
  @shift.setter
  def shift(self, value: Shift):
    self.__shift = value
    if self.research_group and self.research_group.shift != value:
      self.research_group.shift = value
    for section in filter(
      lambda section: section.shift != value,
      self.attending_sections.values()
    ):
      section.shift = value
      
  # Return the available sessions of this student
  @property
  def available_sessions(self):
    sessions = set[Session]()
    if self.shift:
      sessions.update(self.shift.sessions)
      sessions.difference_update(self.attending_sessions)
    return sorted(sessions)
  
# Group abstraction implementation in Python 3.10.6
class ResearchGroup:
  alias: str
  parent: Course
  students: set[Student]
  __shift: Shift | None
  
  # Initialize the instance of this class
  def __init__(self, alias: str, parent: Course):
    self.alias = alias
    self.parent = parent
    self.students = set()
    self.__shift = None
    
  # String representation of any instance of this class
  def __repr__(self):
    return f'{self.parent} {self.alias}'
  

  # This property was implemented instead of simply renaming “__shift” to
  # “shift” due to additional steps done when assigning a value to “shift”
  @property
  def shift(self):
    return self.__shift
  
  # Every time the “shift” value is changed, also change the affected
  # students’ shifts in the case that they differ from the assigned 
  # value.
  @shift.setter
  def shift(self, value: Shift):
    self.__shift = value
    for student in filter(
      lambda student: student.shift != value,
      self.students
    ):
      student.shift = value
      
  # Return the available sessions of the students in this group
  @property
  def available_sessions(self):
    sessions = set[Session]()
    if self.shift:
      sessions.update(self.shift.sessions)
    for student in self.students:
      sessions.difference_update(student.attending_sessions)
    return sorted(sessions)
  
  # Add a student to this group
  def add(self, student: Student):
    student.research_group = self
    self.students.add(student)
    if self.shift:
      student.shift = self.shift
    elif student.shift:
      self.shift = student.shift
  
# Section abstraction implementation in Python 3.10.6
class Section:
  parent: Course
  capacity: Capacity
  students: set[Student]
  parallel_session: ParallelSession | None
  __shift: Shift | None
  

  # Initialize the instance of this class
  def __init__(
    self, 
    parent: Course, 
    shift: Shift | None, 
    parallel_session: ParallelSession | None
  ):
    self.parent = parent
    self.capacity = parent.capacity_section
    self.students = set()
    self.parallel_session = parallel_session
    self.__shift = shift
    
  # String representation of any instance of this class
  def __repr__(self):
    return '{} {}'.format(
      self.parent, 
      self.parallel_session or self.shift
    )
    
  # This property was implemented instead of simply renaming “__shift” to
  # “shift” due to additional steps done when assigning a value to “shift”
  @property
  def shift(self):
    return self.__shift
  
  # Return the available sessions of the students in this section
  @shift.setter
  def shift(self, value: Shift):
    self.__shift = value
    for student in filter(
      lambda student: student.shift != value,
      self.students
    ):
      student.shift = value
      
  # A student is only qualified to to part of a section if and only if:
  # 1. The student qualifies to the course
  # 2. The student has no shift, or the same shift as the section
  # 3. The student is not taking the session of this section
  def qualified(self, student: Student):
    return all([
      self.parent.qualified(student),
      student.shift in {self.shift or student.shift, None},
      self.parallel_session.session in student.available_sessions
        if self.parallel_session and student.shift else True
    ])
    

  # Add a student to this section
  def add(self, student: Student, course_type: CourseType):
    if self.qualified(student):
      if self.shift:
        student.shift = self.shift
      elif student.shift:
        self.shift = student.shift
      self.students.add(student)
      student.courses_taking[course_type] = self.parent
      student.attending_sections[self.parent] = self
      if self.parallel_session:
        student.attending_sessions.add(
          self.parallel_session.session
        )
      return True
    return False
  
  # Add a student to this section, regardless of the number of students
  # attending this section
  def overload(self, student: Student, course_type: CourseType):
    return self.add(student, course_type)
  
  # Add a student to this section—rejecting the student if the number of
  # students exceed or equal the maximum capacity of this section
  def enroll(self, student: Student, course_type: CourseType):
    if len(self.students) < self.capacity.maximum:
      return self.add(student, course_type)
    return False
  
# Course abstraction implementation in Python 3.10.6
class Course:
  alias: str
  difficulty_level: int | None
  capacity_section: Capacity
  capacity_sections: Capacity
  linked_to: Course | None
  sections: set[Section]
  not_alongside: set[Course]
  prerequisites: set[Iterable[Course]]
  
  # Initialize the instance of this class
  def __init__(self, alias: str, difficulty_level: int):
    self.alias = alias
    self.difficulty_level = difficulty_level
    self.linked_to = None
    self.capacity_section = None  # type: ignore
    self.capacity_sections = None  # type: ignore
    self.sections = set()
    self.not_alongside = set([self])
    self.prerequisites = set()
    

  # String representation of any instance of this class
  def __repr__(self):
    return '{}{}'.format(
      self.alias,
      f' Level {self.difficulty_level}'
        if self.difficulty_level else ''
    )
    
  # Return whether this course could open a section
  @property
  def could_open_section(self):
    courses = list(
      course for course in [self, self.linked_to] if course
    )
    return sum(
      len(course.sections) for course in courses
    ) < self.capacity_sections.maximum
  
  # Add a course that shouldn’t be taken alongside this course
  def add_not_alongside(self, not_alongside: Course):
    self.not_alongside.add(not_alongside)
    
  # Add a course that is a prerequisite of this course 
  def add_prerequisites(self, prerequisites: Iterable[Course]):
    self.prerequisites.add(prerequisites)
    
  # Return the sections of this course where the “value” could be found
  def list_sections_by(self, value: Session | Shift):
    result = list[Section]()
    for section in self.sections:
      if value in {
        section.shift,
        section.parallel_session.session
          if section.parallel_session else value
      }:
        result.append(section)
    return result
  
  # A student is qualified to take a course if and only if:
  # 1. The student has not taken this course before
  # 2. The student is not taking a course that shouldn’t be taken alongside
  #    this section
  # 3. The student has taken all prerequisite courses of this course
  def qualified(self, student: Student):
    return all([
      self not in student.courses_taken,
      not self.not_alongside.intersection(student.courses_taking),
      all(
        set(prerequisites).intersection(student.courses_taken)
          for prerequisites in self.prerequisites
      )
    ])
    
  # Overload a student to a section of this course. If the course has 
  # shifted sections, and any of them could accept the student without 
  # being overloaded, the section with fewest students will be prioritized. 
  # Shifted sections are overloaded if ever there aren’t any nonshifted 
  # sections
  def overload(self, student: Student, course_type: CourseType):
    shifted = list(filter(
      lambda section: section.shift and section.qualified(student),
      self.sections
    ))
    noshift = list(filter(
      lambda section: not section.shift and section.qualified(student),
      self.sections
    ))
    if noshift:
      for section in filter(
        lambda section: len(section.students) >= section.capacity.maximum,
        list(shifted)
      ):
        shifted.pop(shifted.index(section))
    if shifted:
      return min(
        shifted,
        key=lambda section: len(section.students)
      ).overload(student, course_type)
    if noshift:
      return min(
        noshift,
        key=lambda section: len(section.students)
      ).overload(student, course_type)
    return False
  # Enroll a student to this course. Unlike the previous function, 
  # this function takes into consideration the capacities of sections
  def enroll(self, student: Student, course_type: CourseType):
    shifted = list(filter(
      lambda section: section.shift and section.qualified(student),
      self.sections
    ))
    noshift = list(filter(
      lambda section: not section.shift and section.qualified(student),
      self.sections
    ))
    for section in shifted + noshift:
      if section.enroll(student, course_type):
        return True
    return False
