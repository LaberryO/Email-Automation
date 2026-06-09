from app.worker import FileLoader
from app.base import BaseUI
from app.defs import ScreenType


class Navigator:
    def __init__(self, loader: FileLoader, ui: BaseUI):
        self.loader = loader
        self.ui = ui

    def handle_input(self):
        text = self.ui.get_input()

        # 종료 신호
        if text == "0":
            return False
        
        if self.ui.screen == ScreenType.MAIN:
            if text == "1":
                self.ui.change(ScreenType.LOAD_DATA)
            if text == "2":
                self.ui.change(ScreenType.SEND_EMAIL)

        return True