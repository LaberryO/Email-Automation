# ui.py
from abc import ABC, abstractmethod
from app.defs import MenuRegistry, Menu

class BaseUI(ABC):
    def __init__(self):
        self._screen: Menu = None

    @property
    def screen(self):
        return self._screen
    
    @screen.setter
    def screen(self, screen: Menu):
        self._screen = screen

    @abstractmethod
    def show(self):
        """UI 출력"""
        pass

    @abstractmethod
    def get_input(self) -> str:
        """input 호출 및 input 값 반환"""
        pass