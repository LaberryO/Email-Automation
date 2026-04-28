import pandas as pd, smtplib, sys, json, time, logging
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

# logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(f"logs/log_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.txt", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# settings
EMAIL_CONTENT = """
<html>
    <body>
        <img src="cid:image" width="600" height="800">
    </body>
</html>
"""

path = "localuser/"

class UserCancelException(Exception): pass

class EmailSender:
    # 파일 로딩
    def __init__(self):
        self.server = None

    def load(self):
        try:
            current_file = "config"
            with open(f"{path}config.json", "r", encoding="utf-8") as file:
                self.config = json.load(file)
                logging.info("config load complete")

            current_file = "image"
            with open(f"{path}image.png", "rb") as file:
                self.image = file.read()
                logging.info("image load complete")

            current_file = "catalog"
            with open(f"{path}catalog.pdf", "rb") as file:
                self.pdf = file.read()
                logging.info("pdf load complete")

        except FileNotFoundError:
            logging.error(f"{current_file} file not found")
            return False
        except PermissionError:
            logging.error(f"{current_file} has no permission")
            return False
        except Exception as e:
            logging.error(f"{current_file} load failed: {e}")
            return False

    # smtp connection
    def connect(self):
        addr = self.config["smtp_address"]
        port = self.config["smtp_port"]
        email = self.config["email"]
        pw = self.config["app_password"]

        try:
            logging.info("try connect smtp server..")
            if self.config.get("SMTP_PORT") == 465:
                logging.info("SMTP PORT is SSL")
                self.server = smtplib.SMTP_SSL(addr, port)
            else:
                logging.info("SMTP PORT is not SSL")
                self.server = smtplib.SMTP(addr, port)
                self.server.starttls()
            self.server.login(email, pw)
            logging.info("stmp connected")

        except smtplib.SMTPAuthenticationError:
            logging.error("incorrect id or password")
            return False
        except Exception as e:
            logging.error(f"smtp connection error: {e}")
            return False
    
    # send email
    def send(self):
        try:
            df = pd.read_csv(f"localuser/data.csv", usecols=self.config["target_cols"])

            # 발송 여부 확인
            check = str(input(f"{len(df)}명에게 정말로 발송하시겠습니까? (Y/N)"))
            if check.lower() == "n":
                raise UserCancelException
            
            for row in df.itertuples():
                user_email = row.이메일
                user_name = row.업체명

                try:
                    msg = MIMEMultipart()
                    msg["Subject"] = self.config["email_subject"]
                    msg["From"] = self.config["email"]
                    msg["To"] = user_email

                    msg.attach(MIMEText(EMAIL_CONTENT, "html"))

                    # 본문에 이미지 넣기
                    image_part = MIMEImage(self.image)
                    image_part.add_header("Content-ID", "<image>")
                    msg.attach(image_part)

                    # PDF 파일 첨부
                    pdf_part = MIMEApplication(self.pdf, _subtype="pdf")
                    pdf_part.add_header("Content-Disposition", "attachment", filename=self.config["pdf_filename"])
                    msg.attach(pdf_part)

                    self.server.send_message(msg)
                    logging.info(f"send email to {user_email}({user_name})")

                except Exception as e:
                    logging.error(f"failed send email: {e}")
                    continue

                finally:
                    logging.info("wait a second..")
                    time.sleep(1.5)

        # 보통 csv 에러 감지
        except Exception as e:
            logging.error(f"send error: {e}")
    
    def close(self):
        if self.server:
            self.server.quit()
            logging.info("server disconect.")

if __name__ == "__main__":
    try:
        app = EmailSender()
        if not app.load():
            raise RuntimeError("data load failed.")
        if not app.connect():
            raise ConnectionError("stmp connect failed.")
        app.send()
        
    except UserCancelException:
        logging.info("user canceled.")
    except Exception as e:
        logging.critical(f"fatal error detected: {e}")
        sys.exit()

    finally:
        app.close()