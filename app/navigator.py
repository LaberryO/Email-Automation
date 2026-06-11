from app.worker import FileLoader
from app.base import BaseUI
from app.defs import MenuRegistry

class Navigator:
    def __init__(self, loader: FileLoader, ui: BaseUI, registry_cls: type[MenuRegistry]):
        self._loader = loader
        self._ui = ui
        self._registry = registry_cls
        self._state: str = MenuRegistry.ROOT.name
        self._is_running: bool = True
    
    @property
    def state(self):
        return self._state

    @property
    def is_running(self):
        return self._is_running

    def handle_input(self):
        """input 호출 및 상태 변경"""
        text = self._ui.get_input()

        current_menu = self._registry.find_by_name(self.state)
        if not current_menu:
            return

        # 0 입력시: 종료 또는 뒤로가기
        if text == "0":
            if current_menu == self._registry.ROOT:
                self._state = self._registry.QUIT.name
                self._is_running = False
            else:
                self._state = self._registry.get_parent_name(self.state) if current_menu.parent else self._registry.ROOT.name
            return
        
        # 숫자 입력시
        try:
            choice_index = int(text) - 1
            if 0 <= choice_index < len(current_menu.children):
                self._state = current_menu.children[choice_index].name
        except (ValueError, TypeError):
            pass
    
    def update(self):
        """비즈니스 로직 및 상태 업데이트"""
        pass

    def render(self):
        """화면 렌더링"""
        current_menu = self._registry.find_by_name(self.state)
        if current_menu:
            self._ui.screen = current_menu
            self._ui.show()