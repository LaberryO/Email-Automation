from abc import ABC, abstractmethod

class BaseEntity:
    def __init__(self, id: int):
        self._id = id
    
    @property
    def id(self):
        return self._id