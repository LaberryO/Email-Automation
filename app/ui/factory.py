# factory.py
from app.base import BaseUI
from app.ui import CLI, GUI

class UIFactory:
    @staticmethod
    def get_ui(mode: str) -> BaseUI:
        if mode == "cli":
            return CLI()
        elif mode == "gui":
            return GUI()
        else:
            raise ValueError("알 수 없는 UI 모드")