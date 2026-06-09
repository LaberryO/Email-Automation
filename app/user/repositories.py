from typing import List
from sqlite3 import Connection

from base import BaseRepository
from user import User

class UserRepository(BaseRepository[User]):
    def __init__(self, conn: Connection):
        # 부모 클래스에 User Entity 주입
        super().__init__(conn, User)