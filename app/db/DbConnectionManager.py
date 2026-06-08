from .DbConfig import DbConfig
from base import BaseConnectionManager as BCM
import sqlite3

class DbConnectionManager(BCM[DbConfig, sqlite3.Connection]):
    def __init__(self, config: DbConfig):
        super().__init__(config)

    def connect(self) -> sqlite3.Connection:
        """SQLite DB에 연결된 Connection 객체 반환"""
        conn = sqlite3.connect(self.config.path)
        conn.row_factory = sqlite3.Row
        return conn