# These imports are needed to have the Allocation Agent run
from src.cls import *
from math import ceil
from src.gnt.EncodeAgent import EncodeAgent

import random

## These constants are for utility purposes
RESEARCH = 'Research'
CORE = 'Core science'
ELEC = 'STE elective'
MATH = 'Math'

# Allocation agent implementation in Python 3.10.6
class AllocateAgent:
  shifts: list[Shift]
  courses: list[Course]
  course_types: dict[str, CourseType]
  __research_groups: list[ResearchGroup]
  __students: list[Student]
  __grouped_students: list[Student]
  __nogroup_students: list[Student]
  
  # Initialize the instance of this class
  def __init__(self, encode_agent: EncodeAgent):
    self.course_types = encode_agent.course_types
    self.shifts = list(encode_agent.shifts)
    self.courses = list(encode_agent.courses.values())
    self.__students = list()
    
    research_groups = set[ResearchGroup]()
    grouped_students = set[Student]()
    nogroup_students = set[Student]()
    for students in encode_agent.students.values():
      self.__students.extend(students)
      for student in students:
        if student.research_group:
          research_groups.add(student.research_group)
          grouped_students.add(student)
        else:
          nogroup_students.add(student)
    self.__research_groups = list(research_groups)
    self.__grouped_students = list(grouped_students)
    self.__nogroup_students = list(nogroup_students)

  # Shuffle the __students field everytime students in accessed
  @property
  def students(self):
    random.shuffle(self.__students)
    return self.__students

  # Shuffle the __research_groups field everytime 
  # research_groups is accessed
  @property
  def research_groups(self):
    random.shuffle(self.__research_groups)
    return self.__research_groups

  # Shuffle the __grouped_students field everytime
  # grouped_students is accessed
  @property
  def grouped_students(self):
    random.shuffle(self.__grouped_students)
    return self.__grouped_students

  # Shuffle the __nogroup_students field everytime
  # nogroup_students is accessed
  @property
  def nogroup_students(self):
    random.shuffle(self.__nogroup_students)
    return self.__nogroup_students
  
  # Open mathematics level sections based on demand
  def open_math_sections(self, course_type: CourseType):
    demand = defaultdict[Course, int](int)
    for student in self.students:
      try:
        demand[student.rankings.current(course_type)] += 1
      except:
        ...
      
    sections_to_open = dict(
      (course, ceil(demand[course] / course.capacity_section.maximum))
        for course in demand
    )
    for course in sections_to_open:
      if course.linked_to:
        while (
          sum(sections_to_open[c] for c in [course, course.linked_to])
            > course.capacity_sections.maximum
        ):
          minimum = course if (
            demand[course] % course.capacity_section.maximum
              < demand[course.linked_to] % course.capacity_section.maximum
          ) else course.linked_to
          sections_to_open[minimum] -= 1
      else:
        sections_to_open[course] = min(
          sections_to_open[course],
          course.capacity_sections.maximum
        )
    for course, amount in sections_to_open.items():
      for _ in range(amount):
        section = Section(course, None, None)
        course.sections.add(section)
        
  # Return the ParallelSession for which the course has available
  def get_parallel_session(self, course: Course, shift: Shift, session: Session):
    return ParallelSession(
      shift, 
      session, 
      len(course.list_sections_by(session))
    )
    
  # Try to section a student according to the highest ranking course
  # that they qualified for
  def try_section_student(
    self, 
    student: Student, 
    course_type: CourseType,
    shift: Shift,
    available_sessions: list[Session]
  ):
    course = student.rankings.current(course_type)
    if course.qualified(student):
      for session in available_sessions:
        for section in course.list_sections_by(session):
          if section.enroll(student, course_type):
            return True
      return False
    student.rankings.final.pop(course_type, 'Not qualified', 0)
    return self.try_section_student(
      student, 
      course_type, 
      shift, 
      available_sessions
    )
  
  # Section a student accourding to the highest ranking course that
  # they are qualified for. This differs from the previous function
  # in the way that this function is allowed to create a new section
  # if the course allows it
  def section_student(
    self, 
    student: Student, 
    course_type: CourseType,
    shift: Shift,
    available_sessions: list[Session]
  ):
    course = student.rankings.current(course_type)
    if course.qualified(student):
      for session in available_sessions:
        for section in course.list_sections_by(session):
          if section.enroll(student, course_type):
            return
      if course.could_open_section:
        section = Section(
          course,
          shift,
          self.get_parallel_session(
            course, 
            shift,
            random.choice(available_sessions)
          )
        )
        course.sections.add(section)
        assert section.enroll(student, course_type)
        return
      student.rankings.final.pop(course_type, 'No rooms available', 0)
    else:
      student.rankings.final.pop(course_type, 'Not qualified', 0)
    self.section_student(student, course_type, shift, available_sessions)
  
  # Allocate students with research groups to courses
  def section_grouped(self):
    for research_group in self.research_groups:
      if not research_group.parent.sections:
        for shift in self.shifts:
          for session in shift.sessions:
            research_group.parent.sections.add(Section(
              research_group.parent,
              shift,
              self.get_parallel_session(
                research_group.parent, 
                shift, 
                session
              )
            ))

      demand = defaultdict[Course, int](int)
      for student in research_group.students:
        core = student.rankings.current(self.course_types[CORE])
        elec = student.rankings.current(self.course_types[ELEC])
        while elec in core.not_alongside:
          student.rankings.final.pop(
            self.course_types[ELEC],
            'Could not be taken alongside Core science',
            0
          )
          elec = student.rankings.current(self.course_types[ELEC])
        demand[core] += 1
        demand[elec] += 1
      
      section_combinations = defaultdict[
        Shift,
        set[tuple[Student, Section, Section]]
      ](set)
      for student in research_group.students:
        core = student.rankings.current(self.course_types[CORE])
        elec = student.rankings.current(self.course_types[ELEC])
        for csection in filter(
          lambda s: 
            len(s.students) + demand[core] <= s.capacity.maximum,
          core.sections
        ):
          for esection in filter(
            lambda s:
              len(s.students) + demand[elec] <= s.capacity.maximum,
              elec.sections
          ):
            if all([
              csection.shift == esection.shift,
              csection.parallel_session.session 
                != esection.parallel_session.session
                if csection.parallel_session
                  and esection.parallel_session
                  else True,
            ]) and csection.shift:
              section_combinations[csection.shift].add(
                (student, csection, esection)
              )
      
      section_excluded_sessions = defaultdict[
        tuple[Shift, Session],
        set[tuple[Student, Section, Section]]
      ](set)
      for shift in section_combinations:
        for session in shift.sessions:
          for student, csection, esection in section_combinations[shift]:
            if (
              csection.parallel_session 
              and csection.parallel_session.session != session
              and esection.parallel_session
              and esection.parallel_session.session != session
            ):
              section_excluded_sessions[(shift, session)].add(
                (student, csection, esection)
              )
      
      research_sections = list(filter(
        lambda section:
          len(section.students) + len(research_group.students)
            <= section.capacity.maximum,
        research_group.parent.sections
      ))
      if section_excluded_sessions:
        for research_section in research_sections:
          shift = research_section.shift
          session = None
          if research_section.parallel_session:
            session = research_section.parallel_session.session
          if (shift, session) not in section_excluded_sessions:
            continue
          for student in research_group.students:
            research_section.enroll(student, self.course_types[RESEARCH])
            if not student.rankings.current(
              self.course_types[MATH]
            ).enroll(
              student,
              self.course_types[MATH]
            ):
              student.rankings.current(self.course_types[MATH]).overload(
                student,
                self.course_types[MATH]
              )
          assert shift and session
          for student, csection, esection in section_excluded_sessions[
            (shift, session)
          ]:
            if not set(student.courses_taking).intersection({
              self.course_types[CORE],
              self.course_types[ELEC]
            }):
              csection.enroll(student, self.course_types[CORE])
              esection.enroll(student, self.course_types[ELEC])
          for student in filter(
            lambda student:
              not set(student.courses_taking).intersection({
                self.course_types[CORE],
                self.course_types[ELEC]
              }),
            research_group.students
          ):
            assert student.shift
            for course_type in [
              self.course_types[CORE], 
              self.course_types[ELEC]
            ]:
              if not self.try_section_student(
                student,
                course_type,
                student.shift,
                student.available_sessions
              ):
                self.section_student(
                  student,
                  course_type,
                  student.shift,
                  student.available_sessions
                )
          break
        if not research_group.shift:
          assert research_group.parent.could_open_section
          shift, session = max(
            section_excluded_sessions,
            key=lambda x: len(section_excluded_sessions[x])
          )
          research_section = Section(
            research_group.parent,
            shift,
            self.get_parallel_session(
              research_group.parent, 
              shift, 
              session
            )
          )
          research_group.parent.sections.add(research_section)
          for student in research_group.students:
            research_section.enroll(student, self.course_types[RESEARCH])
            if not student.rankings.current(
              self.course_types[MATH]
            ).enroll(
              student, 
              self.course_types[MATH]
            ):
              student.rankings.current(self.course_types[MATH]).overload(
                student,
                self.course_types[MATH]
              )
          for student, csection, esection in section_excluded_sessions[
            (shift, session)
          ]:
            if not set(student.courses_taking).intersection({
              self.course_types[CORE],
              self.course_types[ELEC]
            }):
              csection.enroll(student, self.course_types[CORE])
              esection.enroll(student, self.course_types[ELEC])
          for student in filter(
            lambda student:
              not set(student.courses_taking).intersection({
                self.course_types[CORE],
                self.course_types[ELEC]
              }),
            research_group.students
          ):
            assert student.shift
            for course_type in [
              self.course_types[CORE], 
              self.course_types[ELEC]
            ]:
              if not self.try_section_student(
                student,
                course_type,
                student.shift,
                student.available_sessions
              ):
                self.section_student(
                  student,
                  course_type,
                  student.shift,
                  student.available_sessions
                )
      else:
        if (
          not research_sections 
          and research_group.parent.could_open_section
        ):
          shift = random.choice(self.shifts)
          research_section = Section(
            research_group.parent,
            shift,
            self.get_parallel_session(
              research_group.parent,
              shift,
              random.choice(list(shift.sessions))
            )
          )
          research_group.parent.sections.add(research_section)
          research_sections.append(research_section)
        research_section = random.choice(research_sections)
        for student in research_group.students:
          assert research_section.enroll(student, self.course_types[RESEARCH])
          if not student.rankings.current(self.course_types[MATH]).enroll(
            student, 
            self.course_types[MATH]
          ):
            student.rankings.current(self.course_types[MATH]).overload(
              student,
              self.course_types[MATH]
            )
          assert student.shift
          for course_type in [self.course_types[CORE], self.course_types[ELEC]]:
            if not self.try_section_student(
              student,
              course_type,
              student.shift,
              student.available_sessions
            ):
              self.section_student(
                student,
                course_type,
                student.shift,
                student.available_sessions
              )
  
  # Allocate students without sections 
  def section_nogroup(self):
    # Find a tuple or core science elective, science and technology
    # elective, and research course sections that do not conflict with
    # eachother, and then allocate the student to a mathematics section
    def enroll_initial(student: Student):
      research = random.choice(list(student.grade_level.courses[
        (self.course_types[RESEARCH], False)
      ]))
      core = student.rankings.current(self.course_types[CORE])
      elec = student.rankings.current(self.course_types[ELEC])
      while elec in core.not_alongside:
        student.rankings.final.pop(
          self.course_types[ELEC],
          'Could not be alongside Core science',
          0
        )
        elec = student.rankings.current(self.course_types[ELEC])
      for shift in self.shifts:
        for c in core.list_sections_by(shift):
          for e in elec.list_sections_by(shift):
            for r in research.list_sections_by(shift):
              if all([
                c.parallel_session and e.parallel_session and
                  c.parallel_session.session 
                  != e.parallel_session.session,
                c.parallel_session and r.parallel_session and
                  c.parallel_session.session 
                  != r.parallel_session.session,
                e.parallel_session and r.parallel_session and
                  e.parallel_session.session 
                  != r.parallel_session.session,
                len(c.students) < c.capacity.maximum,
                len(e.students) < e.capacity.maximum,
                len(r.students) < r.capacity.maximum
              ]):
                assert (
                  c.enroll(student, self.course_types[CORE]) 
                    and e.enroll(student, self.course_types[ELEC]) 
                    and r.enroll(student, self.course_types[RESEARCH])
                )
                if not student.rankings.current(
                  self.course_types[MATH]
                ).enroll(student, self.course_types[MATH]):
                  student.rankings.current(
                    self.course_types[MATH]
                  ).overload(student, self.course_types[MATH])
                return True
      return False
                
    # Force a student to be allocated to any two core science elective
    # and science and technology elective courses now that there is too
    # few demand to open any other section
    def enroll_final(student: Student):
      for core in filter(
        lambda x: x.qualified(student), 
        student.grade_level.courses[(self.course_types[CORE], True)]
      ):
        for elec in filter(
          lambda x:
            (
              x != core 
              and x not in core.not_alongside 
              and x.qualified(student)
            ), 
          student.grade_level.courses[(self.course_types[ELEC], True)]
        ):
          for c in core.sections:
            for e in elec.sections:
              if all([
                c.parallel_session and
                  c.parallel_session.session in student.available_sessions,
                e.parallel_session and
                  e.parallel_session.session in student.available_sessions,
                c.parallel_session and e.parallel_session and
                  c.parallel_session.session 
                  != e.parallel_session.session,
                len(c.students) < c.capacity.maximum,
                len(e.students) < e.capacity.maximum,
              ]):
                assert (
                  c.enroll(student, self.course_types[CORE]) 
                  and e.enroll(student, self.course_types[ELEC])
                )
                if core != student.rankings.current(self.course_types[CORE]):
                  student.rankings.final.pop(
                    self.course_types[CORE], 
                    'Needed to reshuffle sections', 
                    0
                  )
                if elec != student.rankings.current(self.course_types[ELEC]):
                  student.rankings.final.pop(
                    self.course_types[ELEC], 
                    'Needed to reshuffle sections',
                    0
                  )
                return True
              
    # Find a tuple of "course_type" electives and research courses
    # that do not conflict with each other, and then allocate the
    # students to a mathematics section
    def enroll_type(student: Student, course_type: CourseType):
      research = random.choice(list(student.grade_level.courses[
        (self.course_types[RESEARCH], False)
      ]))            
      course = student.rankings.current(course_type)
      for shift in self.shifts:
        for c in course.list_sections_by(shift):
          for r in research.list_sections_by(shift):
            if all([
              c.parallel_session and r.parallel_session and
                c.parallel_session.session != r.parallel_session.session,
              len(c.students) < c.capacity.maximum,
              len(r.students) < r.capacity.maximum 
                if any(
                  len(s.students) < s.capacity.maximum 
                    for s in research.sections
                ) else True
            ]):
              assert (
                c.enroll(student, course_type) 
                  and (
                    r.enroll(student, self.course_types[RESEARCH]) 
                      or research.overload(
                        student, 
                        self.course_types[RESEARCH]
                      )
                  )
              )
              if not student.rankings.current(
                self.course_types[MATH]
              ).enroll(student, self.course_types[MATH]):
                student.rankings.current(
                  self.course_types[MATH]
                ).overload(student, self.course_types[MATH])
              return True
              
    for student in self.nogroup_students:
      research = random.choice(list(student.grade_level.courses[(
        self.course_types[RESEARCH],
        False
      )]))  
      if not research.sections:
        for shift in self.shifts:
          for session in shift.sessions:
            research.sections.add(Section(
              research, 
              shift, 
              self.get_parallel_session(research, shift, session)
            ))
        for _ in range(2):
          shift = random.choice(self.shifts)
          research_section = Section(
            research, 
            shift, 
            self.get_parallel_session(
              research, 
              shift, 
              random.choice(list(shift.sessions))
            )
          )
          research.sections.add(research_section)
      if (
        not enroll_initial(student) 
          and all(
            not enroll_type(student, type) 
            for type in [self.course_types[CORE], self.course_types[ELEC]]
          )
      ):
        if not research.enroll(student, self.course_types[RESEARCH]):
          research.overload(student, self.course_types[RESEARCH])
        if not student.rankings.current(
          self.course_types[MATH]
        ).enroll(student, self.course_types[MATH]):
          student.rankings.current(
            self.course_types[MATH]
          ).overload(student, self.course_types[MATH])
        assert student.shift
        self.section_student(
          student, 
          self.course_types[CORE], 
          student.shift, 
          student.available_sessions
        )
      for type in filter(
        lambda x: x not in student.courses_taking, 
        [self.course_types[CORE], self.course_types[ELEC]]
      ):
        student.rankings.current(type).enroll(student, type)
    for course in self.courses:
      for section in filter(
        lambda x: len(x.students) < x.capacity.minimum,
        list(course.sections)
      ):
        for student in list(section.students):
          for type, _ in filter(
            lambda x: x[1] == course, 
            list(student.courses_taking.items())
          ):
            student.attending_sections.pop(
              student.courses_taking.pop(type)
            ).students.remove(student)
            if section.parallel_session:
              student.attending_sessions.remove(
                section.parallel_session.session
              )
        course.sections.remove(section)
    for student in self.students:
      for type in filter(
        lambda x: x not in student.courses_taking, 
        [self.course_types[CORE], self.course_types[ELEC]]
      ):
        student.rankings.current(type).enroll(student, type)
        
    for _ in range(10):
      demand = defaultdict[Course, set[tuple[CourseType, Student]]](set)
      for student in self.students:
        for type in filter(
          lambda x: x not in student.courses_taking, 
          [self.course_types[CORE], self.course_types[ELEC]]
        ):
          course = student.rankings.current(type)
          if not course.enroll(student, type):
            demand[course].add((type, student))
      for course, pairs in demand.items():
        sessioned = defaultdict[
          tuple[Shift, Session], 
          set[tuple[CourseType, Student]]
        ](set)
        for type, student in pairs:
          for session in student.available_sessions:
            assert student.shift
            sessioned[(student.shift, session)].add((type, student))
        shift, session = max(sessioned, key=lambda x: len(sessioned[x]))
        if (
          len(sessioned[(shift, session)]) >= course.capacity_section.minimum
        ):
          if course.could_open_section:
            course.sections.add(Section(course, shift, self.get_parallel_session(course, shift, session)))
            for type, student in pairs:
              course.enroll(student, type)
          else:
            for type, student in pairs:
              student.rankings.final.pop(
                type, 
                'Cannot open anymore rooms', 
                0
              )
              if self.course_types[ELEC] not in student.courses_taking:
                core = None
                if self.course_types[CORE] not in student.courses_taking:
                  core = student.rankings.current(self.course_types[CORE])
                else: 
                  core = student.courses_taking[self.course_types[CORE]]
                elec = student.rankings.current(self.course_types[ELEC])
                while core == elec or elec in core.not_alongside:
                  student.rankings.final.pop(
                    self.course_types[ELEC], 
                    'Not compatible with Core science', 
                    0
                  )
                  elec = student.rankings.current(self.course_types[ELEC])
              if (
                self.course_types[CORE] not in student.courses_taking 
                and self.course_types[ELEC] in student.courses_taking
              ):
                elec = student.courses_taking[self.course_types[ELEC]]
                core = student.rankings.current(self.course_types[CORE])
                countdown = 10
                while core == elec or core in elec.not_alongside:
                  student.rankings.final.pop(
                    self.course_types[CORE], 
                    'Not compatible with Core science', 
                    0
                  )
                  core = student.rankings.current(self.course_types[CORE])
                  countdown -= 1
                  if not countdown:
                    assert False
        else:
          for type, student in pairs:
            student.rankings.final.pop(
              type, 
              'Too few demand to open another room', 
              0
            )
            if self.course_types[ELEC] not in student.courses_taking:
              core = None
              if self.course_types[CORE] not in student.courses_taking:
                core = student.rankings.current(self.course_types[CORE])
              else: 
                core = student.courses_taking[self.course_types[CORE]]
              elec = student.rankings.current(self.course_types[ELEC])
              while core == elec or elec in core.not_alongside:
                student.rankings.final.pop(
                  self.course_types[ELEC], 
                  'Not compatible with Core science', 
                  0
                )
                elec = student.rankings.current(self.course_types[ELEC])
            if (
              self.course_types[CORE] not in student.courses_taking 
              and self.course_types[ELEC] in student.courses_taking
            ):
              elec = student.courses_taking[self.course_types[ELEC]]
              core = student.rankings.current(self.course_types[CORE])
              countdown = 10
              while core == elec or core in elec.not_alongside:
                student.rankings.final.pop(
                  self.course_types[CORE], 
                  'Not compatible with Core science', 
                  0
                )
                core = student.rankings.current(self.course_types[CORE])
                countdown -= 1
                if not countdown:
                  assert False
              
    for student in self.students:
      if len(student.courses_taking) != len(self.course_types):
        for type in filter(
          lambda x: x in student.courses_taking,
          [self.course_types[CORE], self.course_types[ELEC]]
        ):
          section = student.attending_sections.pop(
            student.courses_taking.pop(type)
          )
          assert section.parallel_session
          student.attending_sessions.remove(section.parallel_session.session)
          section.students.remove(student)
        assert enroll_final(student)
  
  # The main driver for the allocation agent
  def solve(self):
    self.open_math_sections(self.course_types[MATH])
    self.section_grouped()
    self.section_nogroup()
    assert all(
      len(s.students) >= s.capacity.minimum 
      for category in self.courses 
      for s in category.sections if s.students
    )