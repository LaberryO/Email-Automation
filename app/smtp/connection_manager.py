# connection_manager.py
from app.smtp import SmtpConfig
from app.base import BaseConnectionManager
import smtplib

# SMTP Connection 관리
class SmtpConnectionManager(BaseConnectionManager[SmtpConfig, smtplib.SMTP]):
    """SMTP Connection 매니저"""
    
    def connect(self) -> smtplib.SMTP:
        """Override: SMTP로 연결된 Server Instance 반환"""
        server = smtplib.SMTP(self.config.address, self.config.port)
        server.starttls()
        server.login(self.config.email, self.config.password)
        return server
    
    def _close(self):
        """Override: SMTP Quit 신호 송신"""
        self.connection.quit()