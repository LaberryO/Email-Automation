# model.py
import weakref

class Menu:
    """UI Screen 전환시 사용할 상수의 구조 정의"""
    def __init__(self, name: str, display_name: str, parent: "Menu" = None):
        self._name = name
        self._display_name = display_name
        self._parent = weakref.ref(parent) if parent else None
        self._children: list["Menu"] = []

        # 자식이 부모를 등록
        if self._parent is not None:
            parent_obj = self._parent()
            if parent_obj:
                parent_obj.add_children(self)

    def add_children(self, child: "Menu"):
        # 중복은 개발자 실수임. 중복 예외 처리 없음.
        self._children.append(child)

    @property
    def name(self):
        return self._name

    @property
    def display_name(self):
        return self._display_name
    
    @property
    def parent(self):
        return self._parent() if self._parent else None
    
    @parent.setter
    def parent(self, parent: "Menu"):
        self._parent = weakref.ref(parent) if parent else None
    
    @property
    def children(self):
        return self._children