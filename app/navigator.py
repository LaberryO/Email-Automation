from app.worker import FileLoader
from app.base import BaseUI
from app.defs import ScreenType


class Navigator:
    def __init__(self, loader: FileLoader, ui: BaseUI):
        self._loader = loader
        self._ui = ui
        self._state = ui.screen

    @property
    def ui(self):
        return self._ui
    
    @property
    def state(self):
        return self._state
    
    @state.setter
    def state(self, state: ScreenType):
        self._state = state

    
    @property
    def loader(self):
        return self._loader

    def handle_input(self):
        text = self.ui.get_input()

        # 종료 신호
        if text == "0":
            self.ui.change(ScreenType.EXIT_PROGRAM)
            return False
        
        if self.state == ScreenType.MAIN:
            if text == "1":
                self.ui.change(ScreenType.LOAD_DATA)
            if text == "2":
                self.ui.change(ScreenType.SEND_EMAIL)
            
        self.change_state()
        return True
    
    def change_state(self):
        self.state = self.ui.screen