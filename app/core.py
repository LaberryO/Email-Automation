# core.py
from app import Navigator
from app.worker import FileLoader
from app.ui import CLI
from app.defs import MenuRegistry

class Core:
    def __init__(self):
        loader = FileLoader()
        ui = CLI()
        self.nav = Navigator(loader, ui, MenuRegistry)

    def run(self):
        self.nav.render()
        while self.nav.is_running:
            self.nav.handle_input()
            self.nav.update()
            self.nav.render()