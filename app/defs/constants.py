# constants.py
from app.defs import Menu
from typing import Optional

class MenuRegistry:
    """메뉴 상수 및 조회 로직 관리"""
    # 상수
    ROOT = Menu("ROOT", "메인 메뉴")
    LOAD_FILE = Menu("LOAD_FILE", "파일 업로드", ROOT)
    SEND_EMAIL = Menu("SEND_EMAIL", "이메일 전송", ROOT)
    QUIT = Menu("QUIT", "프로그램 종료")

    # 상수 매핑
    _REGISTRY: dict[str, Menu] = {
        v.name: v for k, v in locals().items() if isinstance(v, Menu)
    }

    @classmethod
    def find_by_name(cls, name: str) -> Optional[Menu]:
        """해당 name을 가진 Menu를 반환합니다."""
        return cls._REGISTRY.get(name)
    
    @classmethod
    def get_parent_name(cls, name: str) -> Optional[str]:
        """해당 name을 가진 상수의 부모 클래스 name을 반환합니다."""
        item = cls.find_by_name(name)
        if item and item.parent:
            return item.parent.name
        return None