from src3 import *
from collections import defaultdict
from src1.gnt.EncodeAgent import EncodeAgent

class AnalyzeAgent:
  @classmethod
  def score(cls, encode_agent: EncodeAgent, show: bool | None = None):
    total = defaultdict[CourseType, int](int)
    attained = defaultdict[CourseType, int](int)
    for students in encode_agent.students.values():
      for student in students:
        for type, _ in filter(
          lambda x: x[1], 
          student.grade_level.courses
        ):
          total[type] += 1
          if type in student.takes and (
            not student.rankings.initial(type) 
            or student.rankings.initial(type) 
              == student.takes[type]
            or student.rankings.final.reason_rejected[
              type][student.rankings.initial(type)]  # type: ignore
              == 'No rooms available'
          ):
            attained[type] += 1
    
    scores = dict((type, attained[type] / total[type] * 100) for type in total)
    if show:
      for type in scores:
        print(f'{scores[type]}% [{type}]')
    return scores