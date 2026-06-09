# entities.py
from app.base import BaseEntity

class Company(BaseEntity):
    def __init__(self, id, name: str, email: str, contact: str):
        super().__init__(id)
        self._name = name
        self._email = email
        self._contact = contact

    @property
    def name(self):
        return self._name
    
    @property
    def email(self):
        return self._email
    
    @property
    def contact(self):
        return self._contact
    
    @classmethod
    def from_data_dict(cls, data, mapping_rules) -> "Company":
        """원본 데이터를 기반으로 Instance 생성"""
        return cls(
            id=data.get("id", 0),
            name=data.get(mapping_rules["name"], ''),
            email=data.get(mapping_rules["email"], ''),
            contact=data.get(mapping_rules["contact"], '')
        )