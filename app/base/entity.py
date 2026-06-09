# entity.py
from abc import ABC, abstractmethod

class BaseEntity(ABC):
    def __init__(self, id: int):
        self._id = id
    
    @property
    def id(self) -> int:
        return self._id
    
    @classmethod
    @abstractmethod
    def from_data_dict(cls, data: dict, mapping_rules: dict) -> "BaseEntity":
        """외부 파일 헤더와 Entity 헤더를 매핑"""
        pass