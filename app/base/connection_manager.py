from abc import ABC, abstractmethod
from typing import TypeVar, Generic, Type
from types import TracebackType
from .config import BaseConfig

T = TypeVar("T", bound=BaseConfig) # Type
C = TypeVar("C") # Connection

class BaseConnectionManager(ABC, Generic[T, C]):
    def __init__(self, config: T):
        self.config: T = config
        self.connection: C | None = None

    @abstractmethod
    def connect(self) -> C:
        """자식 클래스에서 Connection Instance 생성하는 로직"""
        pass

    @abstractmethod
    def _close(self) -> None:
        """자식 클래스에서 Instance Resource 회수하는 로직"""
        pass

    def _commit_or_rollback(self, exc_type: type[BaseException] | None) -> None:
        """선택 의존성: DB Transaction 처리 필요 시 Override"""
        pass

    def __enter__(self) -> C:
        """공통: with 진입 시 Instance 반환"""
        self.connection = self.connect()
        return self.connection
    
    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None
    ) -> bool | None:
        """공통: Try-With-Resource. 자원 해제 로직"""
        if self.connection is not None:
            try:
                self._commit_or_rollback(exc_type)
            except Exception:
                pass
            finally:
                try:
                    self._close()
                except Exception:
                    pass
        return False # 예외는 외부로 Throw