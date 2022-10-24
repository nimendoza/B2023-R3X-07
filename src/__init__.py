from src.gnt.AllocateAgent import AllocateAgent
from src.gnt.AnalyzeAgent import AnalyzeAgent
from src.gnt.EncodeAgent import EncodeAgent
from src.gnt.ReadWriteAgent import ReadWriteAgent
from typing import Any
from time import time

def run():
  encode_agent = EncodeAgent()
  encode_agent.encode_courses('input/Test Data_ Subjects.xlsx')
  encode_agent.encode_students('input/Test Data_ Students.xlsx')

  allocate_agent = AllocateAgent(encode_agent)
  allocate_agent.solve()
  return encode_agent

def export(
  encode_agent: EncodeAgent, 
  directory: str, 
  template: str, 
  time: int
):
  filename = ReadWriteAgent.find_filepath(directory, template)
  data: list[list[Any]] = [
    ['Time taken:', f'{time} s'], 
    ['Course Type', 'Score (%)']
  ]
  for type, score in AnalyzeAgent.score(encode_agent).items():
    data.extend([[type, score]])
  ReadWriteAgent.to_xlsx(filename, 'Summary', data)
  for grade_level, students in encode_agent.students.items():
    types = sorted(encode_agent.course_types.values())
    data = [['Student']]
    data[0].extend(types)
    data[0].extend(map(lambda x: f'Remarks: {x}', types[:3]))
    for student in students:
      line: list[Any] = [student]
      for type in types:
        line.append(
          student.attending_sections[student.courses_taking[type]] 
          if type in student.courses_taking else ''
        )
      for type in types[:3]:
        if student.rankings.initial(type):
          if type in student.courses_taking:
            line.append(
              '{}: {}'.format(
                student.rankings.initial(type),
                student.rankings.final.reason_rejected[type][
                  student.rankings.initial(type)
                ]  # type: ignore
              )
              if student.courses_taking[type] 
              != student.rankings.initial(type) else ''
            )
          else:
            line.append(f'Wants {student.rankings.initial(type)}')
        else:
          line.append('Initial rankings are invalid')
      data.append(line)
    ReadWriteAgent.to_xlsx(filename, str(grade_level), data)
  for category in filter(
    lambda x: x.sections, 
    encode_agent.courses.values()
  ):
    data = list(
      list() 
      for _ in range(max(map(
        lambda x: len(x.students), 
        category.sections)) + 1)
    )
    for section in category.sections:
      data[0].append(section)
      for i, student in enumerate(section.students, 1):
        data[i].append(student)
    ReadWriteAgent.to_xlsx(filename, str(category), data)

def main(target: int):
  initial_time = time()
  encode_agent = None
  while not encode_agent:
    try:
      encode_agent = run()
    except:
        continue
    results = AnalyzeAgent.score(encode_agent, True)
    if min(results.values()) < target or results[
      encode_agent.course_types['Math']
    ] < 100:
      encode_agent = None
      continue
    export(
      encode_agent, 
      'output', 
      f'Test Results {"{}"}: Aiming for {target}%.xlsx', 
      int(time() - initial_time)
    )
  # AnalyzeAgent.score(encode_agent, True)
  # export(encode_agent, 'output', 'Test Results {}.xlsx')

if __name__ == '__main__':
  # encode_agent = run()
  # AnalyzeAgent.score(encode_agent, True)
  main(95)