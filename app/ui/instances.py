# instances.py
from app.base import BaseUI
from app.defs import ScreenType

class CLI(BaseUI):
    def show(self) -> ScreenType:
        self._render(self.screen)

    def change(self, screen: ScreenType):
        self.screen = screen

    def get_input(self, is_text: bool = False) -> str:
        if is_text:
            return input("TEXT: ")
        else:
            return input("> ").strip()
        
    def _render(self, screen: ScreenType):
        header = screen.value[0]
        items = screen.value[1:]

        print(f"\n{header}")
        print("0. 종료")
        for index, item in enumerate(items, 1):
            print(f"{index}. {item}")
        print("")

    
class GUI(BaseUI):
    pass