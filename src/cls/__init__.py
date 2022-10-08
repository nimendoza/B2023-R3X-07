from __future__ import annotations

from collections    import defaultdict
from typing         import Iterable
from uuid           import UUID, uuid4


class Capacity:
    id:      UUID
    ideal:   int
    minimum: int
    maximum: int
    
    def __init__(self, minimum: int, maximum: int, ideal: int = None):
        self.id      = uuid4()
        self.ideal   = ideal or minimum
        self.minimum = minimum
        self.maximum = maximum
        assert 0 <= minimum <= self.ideal <= maximum
        
    def __repr__(self):
        return f'Minimum: {self.minimum} | Ideal: {self.ideal} | Maximum: {self.maximum}'
    
    def __eq__(self, other: Capacity):
        return all([
            self.minimum == other.minimum,
            self.maximum == other.maximum
        ])
        
    def copy(self):
        return Capacity(self.minimum, self.maximum)
    
    
class Partition:
    id:    UUID
    alias: str
    
    def __init__(self, alias: str):
        self.id    = uuid4()
        self.alias = alias
        
    def __repr__(self):
        return self.alias
    
    # def __eq__(self, other: Partition):
    #     return self.alias == other.alias
    
    def __lt__(self, other: Partition):
        return self.alias < other.alias
    
    
class Shift:
    id:         UUID
    partitions: set[Partition]
    
    def __init__(self):
        self.id         = uuid4()
        self.partitions = set()
        
    def __repr__(self):
        return ''.join(map(str, sorted(self.partitions)))
    
    def __contains__(self, partition: Partition):
        return partition in self.partitions
    
    def add(self, partition: Partition):
        self.partitions.add(partition)
        
        
class ParallelSession:
    id:        UUID
    shift:     Shift
    partition: Partition
    index:     int
    
    def __init__(self, shift: Shift, partition: Partition, index: int = None):
        self.id = uuid4()
        assert partition in shift
        self.shift     = shift
        self.partition = partition
        self.index     = index or 0
        
    def __repr__(self):
        return '{}{}'.format(self.partition, self.index or '')
    
    def __lt__(self, other: ParallelSession):
        if self.partition.alias == other.partition.alias:
            return self.index < other.index
        return self.partition < other.partition
    
    
class Type:
    id:    UUID
    alias: str
    order: int
    
    def __init__(self, alias: str, order: int):
        self.id    = uuid4()
        self.alias = alias
        self.order = order
        
    def __repr__(self):
        return self.alias
    
    def __lt__(self, other: Type):
        return self.order < other.order
    
    
class GradeLevel:
    id:         UUID
    alias:      int
    categories: defaultdict[tuple[Type, bool], set[Category]]
    
    def __init__(self, alias: int):
        self.id         = uuid4()
        self.alias      = alias
        self.categories = defaultdict(set)
        
    def __repr__(self):
        return f'Grade {self.alias}'
    
    def __lt__(self, other: GradeLevel):
        return self.alais < other.alias
    
    def add(self, type: Type, ranked: bool, category: Category):
        self.categories[(type, ranked)].add(category)
        
        
class Rank:
    id:          UUID
    alais:       str
    grade_level: GradeLevel
    ordered:     dict[Type, list[Category]]
    reasons:     dict[Type, dict[Category, str]]
    
    def __init__(self, alias: int, grade_level: GradeLevel):
        self.id = uuid4()
        self.alias = alias
        self.grade_level = grade_level
        self.ordered = dict((pair[0], list()) for pair in grade_level.categories if pair[1])
        self.reasons = dict((pair[0], dict()) for pair in grade_level.categories if pair[1])
        
    def __repr__(self):
        return f'Rank {self.alias}'
    
    def len(self, type: Type):
        return len(self.ordered[type])
    
    def add(self, type: Type, category: Category):
        assert category in self.grade_level.categories[(type, True)]
        self.ordered[type].append(category)
        
    def pop(self, type: Type, reason: str, index: int = None):
        self.reasons[type][self.ordered[type].pop(index or 0)] = reason
        
    def get(self, type: Type, index: int = None):
        return self.ordered[type][index or 0]
    
    def copy(self, type: Type):
        return list(self.ordered[type])
    
    
class Rankings:
    id:    UUID
    start: Rank
    final: Rank
    reset: Rank
    owner: Student
    
    def __init__(self, owner: Student):
        self.id    = uuid4()
        self.owner = owner
        self.start = Rank('Start', owner.grade_level)
        self.final = Rank('Final', owner.grade_level)
        self.reset = Rank('Reset', owner.grade_level)
        for type, ranked in filter(lambda x: x[1], owner.grade_level.categories):
            for category in filter(lambda x: x.qualified(owner), owner.grade_level.categories[(type, ranked)]):
                self.reset.add(type, category)
                
    def __repr__(self):
        return f'{self.owner}\'s Rankings'
    
    def add(self, type: Type, category: Category):
        if category.qualified(self.owner):
            self.start.add(type, category)
            self.final.add(type, category)
            
    def initial(self, type: Type, index: int = None):
        if not self.start.len(type):
            return None
        return self.start.get(type, index)
    
    def current(self, type: Type, index: int = None):
        if not self.final.len(type):
            self.final.ordered[type] = self.reset.copy(type)
        return self.final.ordered[type][index or 0]
    
    
class Student:
    __shift: Shift | None
    
    id:          UUID
    alias:       str
    group:       Group | None
    grade_level: GradeLevel
    
    # TODO: Implement these
    prerequisites: set[Student]
    not_alongside: set[Student]
    
    rankings: Rankings
    taken:    set[Category]
    sessions: set[Partition]
    takes:    dict[Type, Category]
    sections: dict[Category, Section]
    
    def __init__(self, alias: str, grade_level: GradeLevel):
        self.__shift = None
        
        self.id          = uuid4()
        self.group       = None
        self.alias       = alias
        self.taken       = set()
        self.takes       = dict()
        self.sessions    = set()
        self.sections    = dict()
        self.grade_level = grade_level
        self.rankings    = Rankings(self)
        
    def __repr__(self):
        return f'{self.grade_level}-{self.alias}'
    
    @property
    def shift(self):
        return self.__shift
    
    @shift.setter
    def shift(self, value: Shift):
        assert self.shift in {value or self.shift, None}
        self.__shift = value
        if self.group and self.group.shift != value:
            self.group.shift = value
        for section in filter(lambda x: x.shift != value, self.sections.values()):
            section.shift = value
            
    @property
    def available_partitions(self):
        partitions = set[Partition]()
        if self.shift:
            partitions.update(self.shift.partitions)
            partitions.difference_update(self.sessions)
        return sorted(partitions)
    
    
class Group:
    __shift: Shift | None
    
    id:       UUID
    alias:    str
    parent:   Category
    students: set[Student]
    
    def __init__(self, alias: str, parent: Category):
        self.__shift = None
        
        self.id       = uuid4()
        self.alias    = alias
        self.parent   = parent
        self.students = set()
        
    def __repr__(self):
        return f'{self.parent} {self.alias}'
    
    @property
    def shift(self):
        return self.__shift
    
    @shift.setter
    def shift(self, value: Shift):
        assert self.shift in {value, None}
        self.__shift = value
        for student in filter(lambda x: x.shift != value, self.students):
            student.shift = value
            
    @property
    def available_partitions(self):
        partitions = set[Partition]()
        if self.shift:
            partitions.update(self.shift.partitions)
        for student in self.students:
            partitions.difference_update(student.sessions)
        return list(partitions)
    
    def add(self, student: Student):
        assert not student.group
        student.group = self
        self.students.add(student)
        if self.shift:
            student.shift = self.shift
        elif student.shift:
            self.shift = student.shift
            
            
class Section:
    __shift: Shift | None
    
    id:       UUID
    parent:   Category
    capacity: Capacity
    students: set[Student]
    session:  ParallelSession | None
    
    def __init__(self, parent: Category, shift: Shift = None, session: ParallelSession = None):
        self.__shift = shift
        
        self.id        = uuid4()
        self.parent    = parent
        self.students  = set()
        self.session   = session
        self.capacity  = parent.capacity_section.copy()
        
    def __repr__(self):
        return '{} {}'.format(self.parent, self.session or self.shift)
    
    @property
    def shift(self):
        return self.__shift
    
    @shift.setter
    def shift(self, value: Shift):
        assert self.shift in {value, None}
        self.__shift = value
        for student in filter(lambda x: x.shift != value, self.students):
            student.shift = value
            
    def qualified(self, student: Student):
        return all([
            self.parent.qualified(student),
            student.shift in {self.shift or student.shift, None},
            self.session.partition not in student.sessions if self.session else True
        ])
        
    def add(self, student: Student, type: Type):
        if self.qualified(student):
            if self.shift:
                student.shift = self.shift
            elif student.shift:
                self.shift = student.shift
            self.students.add(student)
            student.takes[type] = self.parent
            student.sections[self.parent] = self
            if self.session:
                student.sessions.add(self.session.partition)
            return True
        return False
    
    def overload(self, student: Student, type: Type):
        return self.add(student, type)
    
    def enroll(self, student: Student, type: Type):
        if len(self.students) < self.capacity.maximum:
            return self.add(student, type)
        return False
    
    def dropout(self, student: Student, type: Type, permanent: bool = None):
        assert student in self.students
        self.students.remove(student)
        student.sections.pop(student.takes.pop(type))
        if self.session:
            student.sessions.remove(self.session.partition)
        if not student.takes:
            student.shift = None
        if permanent:
            if student.rankings.current(type) == self.parent:
                student.rankings.final.pop(type, 'Dropped out')


class Category:
    id:                UUID
    alias:             str
    level:             int | None
    capacity_section:  Capacity
    capacity_sections: Capacity
    link:              Category | None
    sections:          set[Section]
    not_alongside:     set[Category]
    prerequisites:     set[Iterable[Category]]
    
    def __init__(self, alias: str, level: int = None):
        self.id            = uuid4()
        self.link          = None
        self.alias         = alias
        self.level         = level
        self.sections      = set()
        self.not_alongside = set()
        self.prerequisites = set()
        
    def __repr__(self):
        return '{}{}'.format(self.alias, f' Level {self.level}' if self.level else '')
    
    def add_not_alongside(self, not_alongside: Category):
        assert self != not_alongside
        self.not_alongside.add(not_alongside)
        
    def add_prerequisites(self, prerequisites: Iterable[Category]):
        assert self not in prerequisites
        self.prerequisites.add(prerequisites)
        
    def list_sections_by(self, value: Partition | Shift):
        result = list[Section]()
        for section in self.sections:
            if value in {section.shift, section.session.partition if section.session else value}:
                result.append(section)
        return result
    
    def could_open_section(self):
        categories = list(category for category in [self, self.link] if category)
        assert all(category.capacity_sections == self.capacity_sections for category in categories)
        return sum(len(category.sections) for category in categories) < self.capacity_sections.maximum
    
    def qualified(self, student: Student):
        return all([
            self not in student.taken,
            self not in student.takes.values(),
            not self.not_alongside.intersection(student.takes.values()),
            all(set(prerequisites).intersection(student.taken) for prerequisites in self.prerequisites)
        ])
        
    def overload(self, student: Student, type: Type):
        shifted = list(filter(lambda x: x.shift and x.qualified(student), self.sections))
        noshift = list(filter(lambda x: not x.shift and x.qualified(student), self.sections))
        if noshift:
            for section in filter(lambda x: len(x.students) >= x.capacity.maximum, list(shifted)):
                shifted.pop(shifted.index(section))
        if shifted:
            return min(shifted, key=lambda x: len(x.students)).overload(student, type)
        if noshift:
            return max(noshift, key=lambda x: len(x.students)).overload(student, type)
        return False
    
    def enroll(self, student: Student, type: Type):
        for section in self.sections:
            if section.enroll(student, type):
                return True