from app.base import BaseConfig

class SmtpConfig(BaseConfig):
    def __init__(self, address: str, port: str, email: str, password: str):
        self._address = address
        self._port = port
        self._email = email
        self._password = password

    @property
    def address(self):
        return self._address
    
    @property
    def port(self):
        return self._port
    
    @property
    def email(self):
        return self._email
    
    @property
    def password(self):
        return self._password