from abc import ABC, abstractmethod
from typing import TypeVar, Type, Generic, List, Optional
from sqlite3 import Connection
from .entity import BaseEntity

T = TypeVar("T", bound=BaseEntity)

# Repository
class BaseRepository(ABC, Generic[T]):
    def __init__(self, conn: Connection, model: Type[T]):
        self.conn = conn
        self.model = model

        # Table Name 규칙
        # Entity -> entity -> entities
        self.table_name = model.__name__.lower() + "s"

        # id는 BaseEntity에 있으니 바로 할당, 나머지는 property 통해서 수집
        self.fields = ["id"] + [
            name for name, obj in vars(model).items()
            if isinstance(obj, property)
        ]

    def find_by_id(self, entity_id: int) -> Optional[T]:
        """PK를 통한 단일 조회"""
        query = f"SELECT * FROM {self.table_name} WHERE id = ?"
        cursor = self.conn.cursor()
        cursor.execute(query, (entity_id,))
        row = cursor.fetchone()

        if row is None:
            return None
        
        # Tuple to Entity 이후 반환
        return self.model(*row)
    
    def save_all(self, entities: List[T]) -> None:
        """대량 적재"""
        if not entities:
            return
        
        # 1. Query Build
        columns = ", ".join(self.fields)
        placeholder = ", ".join(["?"] * len(self.fields))
        query = f"INSERT INTO {self.table_name} ({columns}) VALUES ({placeholder})"

        # 2. Raw Value 추룰
        data_to_insert = [
            tuple(getattr(entity, field) for field in self.fields)
            for entity in entities
        ]

        cursor = self.conn.cursor()
        cursor.executemany(query, data_to_insert)