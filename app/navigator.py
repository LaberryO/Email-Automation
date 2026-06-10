from app.worker import FileLoader
from app.base import BaseUI
from app.defs import MenuRegistry

class Navigator:
    def __init__(self, loader: FileLoader, ui: BaseUI, registry_cls: type[MenuRegistry]):
        self._loader = loader
        self._ui = ui
        self._registry = registry_cls
        self._state: str = None
    
    @property
    def state(self):
        return self._state
    
    @state.setter
    def state(self, state: str):
        self._state = state

    def handle_input(self):
        """input 호출 및 상태 전달"""
        # 최초 시작 시 스킵
        if not self.state:
            return

        text = self._ui.get_input()

        current_menu = self._registry.find_by_name(self.state)
        if not current_menu:
            return

        if text == "0":
            if current_menu == self._registry.ROOT:
                self.state = self._registry.QUIT.name
            else:
                self.state = self._registry.get_parent_name(self.state) if current_menu.parent else self._registry.ROOT.name
            return
        
        try:
            choice_index = int(text) - 1
            if 0 <= choice_index < len(current_menu.children):
                self.state = current_menu.children[choice_index].name
        except ValueError:
            pass
    
    def update(self) -> bool:
        """화면 업데이트"""
        if not self.state:
            self.state = self._registry.ROOT.name

        current_menu = self._registry.find_by_name(self.state)

        self.ui.screen = current_menu
        self.ui.show()

        return self.state != self._registry.QUIT.name