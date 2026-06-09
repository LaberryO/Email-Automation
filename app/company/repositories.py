from typing import List
from sqlite3 import Connection

from app.base import BaseRepository
from app.company import Company

class CompanyRepository(BaseRepository[Company]):
    def __init__(self, conn: Connection):
        # 부모 클래스에 Entity 주입
        super().__init__(conn, Company)