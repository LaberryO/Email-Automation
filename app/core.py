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
        running = True
        while running:
            self.nav.handle_input()
            running = self.nav.update()