from __future__  import annotations
from copy        import deepcopy
from datetime    import date
from math        import ceil
from os.path       import exists
from src.classes   import *
from src.utilities import *
from time          import time
from typing        import Any, Optional

import sys

def export(
  data: Data, 
  time_taken: int,
  number_of_guesses: int,
  targets: Optional[dict[CourseType, float]] = None):
  filename = find_filepath(RESULT_FILEPATH, RESULT_FILENAME)
  sheet: list[list[Any]] = [
    ['Time taken (s):', time_taken],
    ['Number of guesses:', number_of_guesses]]
  if targets:
    sheet.append(['Course Type', 'Target (%)', 'Score (%)'])
    sheet += list(
      [course_type, targets[course_type], actual_score]
      for course_type, actual_score in score(data).items())
  else:
    sheet.append(['Course Type', 'Score (%)'])
    sheet += list(
      [course_type, actual_score]
      for course_type, actual_score in score(data).items())
  to_xlsx(filename, 'Summary', sheet)
  for grade_level, students in data.students.items():
    course_types = sorted(data.course_types.values())
    ranked_types = sorted(
      course_type for course_type, ranked in grade_level.courses if ranked)
    sheet = [['Student'] + course_types + list(
      map(lambda course_type: f'Remarks: {course_type}', ranked_types))]
    for student in sorted(students):
      line: list[Any] = [student] + list(
        student.sections[student.takes[course_type]] 
        if course_type in student.takes else ''
        for course_type in course_types )
      for course_type in ranked_types:
        if student.rankings.initial(course_type):
          if course_type in student.takes:
            line.append(
              '{}: {}'.format(
                student.rankings.initial(course_type),
                student.rankings.final.reason_rejected[course_type][
                  student.rankings.initial(course_type)
                ])  # type: ignore
              if student.takes[course_type] != student.rankings.initial(
                course_type) else '')
          else:
            line.append(f'Wants {student.rankings.initial(course_type)}')
        else:
          line.append('Initial rankings were invalid')
      sheet.append(line)
    to_xlsx(filename, str(grade_level), sheet)
  for course in data.courses.values():
    if course.sections:
      sheet = list(list() for _ in range(max(map(
        lambda section: len(section.students), course.sections)) + 1))
      for section in sorted(course.sections):
        sheet[0].append(section)
        for i, student in enumerate(sorted(section.students), 1):
          sheet[i].append(student)
        for j in range(i + 1, len(sheet)):  # type: ignore
          sheet[j].append('')
      to_xlsx(filename, str(course), sheet)
    
def main():
  system_path   = 'input/Test Data_ Subjects.xlsx'
  students_path = 'input/Test Data_ Students.xlsx'
  if not DEBUG:
    system_path   = input('System input file relative path: ')
    students_path = input('Student input file relative path: ')
  if not exists(system_path):
    raise Exception(f'\'{system_path}\' does not exist')
  if not exists(students_path):
    raise Exception(f'\'{students_path} does not exists\'')
  
  data = Data()
  encode(data, system_path, students_path)
  
  print('Please select mode:')
  print('  [1] Given target preferences, get a certain number of results')
  print('  [2] Given a certain number of guesses, get the best result')
  match int(input('Mode chosen: ')):
    case 1:
      target_scores = dict(get_target_scores(data))
      results_count = int(input('Number of results to produce: '))
      for _ in range(results_count):
        initial_time = time()
        number_of_guesses = 0
        while not meets_target_scores(data, target_scores):
          number_of_guesses += 1
          solve(data)
        time_taken = ceil(time() - initial_time)
        export(data, time_taken, number_of_guesses, target_scores)
    case 2:
      initial_time = time()
      best = deepcopy(data)
      guess_count = int(input('Number of iterations to choose from: '))
      for _ in range(guess_count):
        solve(data)
        if sum(score(best).values()) < sum(score(data, True).values()):
          best = deepcopy(data)
      time_taken = ceil(time() - initial_time)
      export(best, time_taken, guess_count)
    case _:
      raise Exception('Mode not supported')

if __name__ == '__main__':
  sys.setrecursionlimit(int(1e9))
  
  RESULT_FILEPATH = 'output'
  RESULT_FILENAME = f'{date.today()} Result {"{}"}.xlsx'
  
  RESEARCH = 'Research'
  MATH = 'Mathematics level'
  CORE = 'Core science elective'
  ELEC = 'Science and technology elective'
  
  DEBUG = True
  
  main()