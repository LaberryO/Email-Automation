from abc import ABC, abstractmethod
from typing import TypeVar, Generic, Type
from .BaseConfig import BaseConfig

T = TypeVar("T", bound=BaseConfig)
C = TypeVar("C")

class BaseConnectionManager(ABC, Generic[T, C]):
    def __init__(self, config: Type[T]):
        self.config = config

    @abstractmethod
    def connect(self) -> Type[C]:
        pass