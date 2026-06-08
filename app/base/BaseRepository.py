from abc import ABC, abstractmethod
from typing import TypeVar, Type, Generic
from sqlite3 import Connection
from .BaseEntity import BaseEntity

T = TypeVar("T", bound=BaseEntity)

# Repository
class BaseRepository(ABC, Generic[T]):
    def __init__(self, conn: Connection, model: Type[T]):
        self.conn = conn
        self.model = model

    def findById(self):
        

