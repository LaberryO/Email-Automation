from db import BaseEntity

class UserEntity(BaseEntity):
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