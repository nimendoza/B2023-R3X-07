# These imports are needed to have this segment of the software run
from datetime import date
from math     import ceil
from src1.gnt.AllocateAgent  import AllocateAgent
from src1.gnt.AnalyzeAgent   import AnalyzeAgent
from src1.gnt.EncodeAgent    import EncodeAgent
from src1.gnt.ReadWriteAgent import ReadWriteAgent
from typing import Any
from time   import time

# Given the file path to the courses input data, ask the user their target
# percentage goal for the output. This will be one of the hard constraints
# of the program
def get_target_values(courses: str):
  encode_agent = EncodeAgent()
  encode_agent.encode_courses(courses)
  
  for course_type, mlem in zip(sorted(encode_agent.course_types.values()), [97, 97, 100, -1]):
    yield (
      course_type, 
      mlem # int(input(f'Target % for {course_type} (-1 for None): '))
    )

# Given the file paths to the course and student input data, as well as 
# the target values for each course type, try to produce a valid output
def run(courses: str, students: str, target_values: dict[str, int]):
  encode_agent = EncodeAgent()
  encode_agent.encode_courses(courses)
  encode_agent.encode_students(students)
  
  allocate_agent = AllocateAgent(encode_agent)
  allocate_agent.solve()
  
  scores_attained = dict(
    (str(course_type), score)
      for course_type, score 
      in AnalyzeAgent.score(encode_agent, True).items()
  )
  for course_type in scores_attained:
    if scores_attained[course_type] < target_values[course_type]:
      raise Exception(f'Target percentage not attained for {course_type}')
  return encode_agent

# Given an instance of the encode agent, write its data to the specified 
# .xlsx file. Also note the time taken to reach that result
def export(
  encode_agent: EncodeAgent,
  output_directory: str,
  target_values: dict[str, int],
  name_template: str,
  time_taken: int,
  number_of_guesses: int
):
  filename = ReadWriteAgent.find_filepath(output_directory, name_template)
  data: list[list[Any]] = [
    ['Time taken (s):', time_taken],
    ['Number of guesses:', number_of_guesses],
    ['Course Type', 'Target (%)', 'Score (%)']
  ] + list(
    [course_type, target_values[str(course_type)], score]
      for course_type, score in AnalyzeAgent.score(encode_agent).items()
  )
  ReadWriteAgent.to_xlsx(filename, 'Summary', data)
  for grade_level, students in encode_agent.students.items():
    course_types = sorted(encode_agent.course_types.values())
    ranked_types = sorted(
      course_type
        for course_type, ranked in grade_level.courses if ranked
    )
    data = [['Student']]
    data[0].extend(course_types)
    data[0].extend(map(lambda x: f'Remarks: {x}', ranked_types))
    for student in sorted(students):
      line: list[Any] = [student]
      for course_type in course_types:
        line.append(
          student.sections[student.takes[course_type]]
            if course_type in student.takes else ''
        )
      for course_type in ranked_types:
        if student.rankings.initial(course_type):
          if course_type in student.takes:
            line.append(
              '{}: {}'.format(
                student.rankings.initial(course_type),
                student.rankings.final.reason_rejected[course_type][
                  student.rankings.initial(course_type)
                ]  # type: ignore
              )
                if student.takes[course_type]
                != student.rankings.initial(course_type) else ''
            )
          else:
            line.append(f'Wants {student.rankings.initial(course_type)}')
        else:
          line.append('Initial rankings were invalid')
      data.append(line)
    ReadWriteAgent.to_xlsx(filename, str(grade_level), data)
  for course in encode_agent.courses.values():
    if course.sections:
      data = list(list() for _ in range(max(map(
        lambda section: len(section.students), course.sections
      )) + 1))
      for section in course.sections:
        data[0].append(section)
        for i, student in enumerate(sorted(section.students), 1):
          data[i].append(student)
        for j in range(i + 1, len(data)):  # type: ignore
          data[j].append('')
      ReadWriteAgent.to_xlsx(filename, str(course), data)

# Driver code implementation in Python 3.10.6
def main():
  course_path =  'input/Test Data_ Subjects.xlsx' # input('Course input file relative path: ')
  student_path = 'input/Test Data_ Students.xlsx' # input('Student input file relative path: ')
  target_values = dict(get_target_values(course_path))
  target_values = dict(
    (str(course_type), target)
      for course_type, target in target_values.items()
  )

  encode_agent = None
  initial_time = time()
  number_of_guesses = 0
  while not encode_agent:
    number_of_guesses += 1
    try:
      encode_agent = run(course_path, student_path, target_values)
    except:
      continue
  export(
    encode_agent, 
    'output/Version 3/97', 
    target_values,
    f'{date.today()} Result {"{}"}.xlsx', 
    ceil(time() - initial_time),
    number_of_guesses
  )

if __name__ == '__main__':
  for mlerm in range(5):
    main()