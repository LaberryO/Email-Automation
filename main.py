# main.py
from app.core import run
from app.ui import CLI
from app import Navigator
from app.worker import FileLoader

if __name__ == "__main__":
    loader = FileLoader()
    ui = CLI()
    nav = Navigator(loader, ui)
    run(ui, nav)