from .SmtpConfig import SmtpConfig
import smtplib
from base import BaseConnectionManager as BCM

# SMTP Connection 관리
class SmtpConnectionManager(BCM[SmtpConfig, smtplib.SMTP]):
    def __init__(self, config: SmtpConfig):
        super().__init__(config)
    
    def connect(self) -> smtplib.SMTP:
        """SMTP로 연결된 Server 반환"""
        server = smtplib.SMTP(self.config.address, self.config.port)
        server.starttls()
        server.login(self.config.email, self.config.password)
        return server