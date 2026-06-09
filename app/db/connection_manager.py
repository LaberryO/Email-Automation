# connection_manager.py
from app.db import DbConfig
from base import BaseConnectionManager
import sqlite3

class DbConnectionManager(BaseConnectionManager[DbConfig, sqlite3.Connection]):
    """DB Connection 매니저"""

    def connect(self) -> sqlite3.Connection:
        """Override: SQLite DB 파일 연결 후 Row Factory 설정된 Instance 반환"""
        conn = sqlite3.connect(self.config.path)
        conn.row_factory = sqlite3.Row
        return conn
    
    def _close(self) -> None:
        """Override: DB Connection Close"""
        self.connection.close()

    def _commit_or_rollback(self, exc_type: type[BaseException] | None) -> None:
        """Override: DB Transaction 제어"""
        if exc_type is None:
            # 에러 없음
            self.connection.commit()
        else:
            # 에러 발생
            self.connection.rollback()