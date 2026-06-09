# core.py
from app.base import BaseUI
from app import Navigator

def run(ui: BaseUI, nav: Navigator):
    while True:
        ui.show()

        if not nav.handle_input():
            ui.show()
            break