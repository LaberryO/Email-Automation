# config.py
from app.base import BaseConfig

class DbConfig(BaseConfig):
    def __init__(self, path: str):
        self._path = path

    @property
    def path(self):
        return self._path