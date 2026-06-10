# instances.py
from app.base import BaseUI
from app.defs import Menu, MenuRegistry

class CLI(BaseUI):
    def show(self):
        self._render(self.screen)

    def get_input(self, is_text: bool = False) -> str:
        if is_text:
            return input("TEXT: ")
        else:
            return input("> ").strip()
        
    def _render(self, screen: Menu):
        header = screen.display_name
        menu = screen.children

        match screen.name:
            case "ROOT":
                zero = "\n0. 종료\n"
            case "QUIT":
                zero = ""
            case _:
                zero = "\n0. 뒤로 가기\n"

        print(f"\n--- {header} ---")
        for index, item in enumerate(menu):
            print(f"{index + 1}. {item.display_name}")
        print(zero)

    
class GUI(BaseUI):
    pass