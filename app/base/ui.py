# ui.py
from abc import ABC, abstractmethod
from app.defs import ScreenType

class BaseUI(ABC):
    def __init__(self):
        self.screen = ScreenType.MAIN

    @abstractmethod
    def show(self):
        pass

    @abstractmethod
    def change(self, screen: ScreenType):
        pass

    @abstractmethod
    def get_input(self) -> str:
        pass