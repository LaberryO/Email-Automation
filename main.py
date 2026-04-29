import pandas as pd, smtplib, sys, json, time, logging, os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

# logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("logs/latest.txt", mode="w", encoding="utf-8"),
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

path = "settings/"

class UserCancelException(Exception): pass

class EmailSender:
    # 파일 로딩
    def __init__(self):
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
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

            return True

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
            logging.info(f"try connect smtp server({port})..")
            if port == 465:
                logging.info("SMTP PORT is SSL")
                self.server = smtplib.SMTP_SSL(addr, port)
            else:
                logging.info("SMTP PORT is not SSL")
                self.server = smtplib.SMTP(addr, port)
                self.server.starttls()
            self.server.login(email, pw)
            logging.info("stmp connected")

            return True

        except smtplib.SMTPAuthenticationError:
            logging.error("incorrect id or password")
            return False
        except Exception as e:
            logging.error(f"smtp connection error: {e}")
            return False
    
    # send email
    def send(self):
        try:
            # 혹시 모를 debug code
            if self.config["debug_mode"]: 
                filename = "debug"
                logging.info(f"debug mode enabled: {path}{filename}.csv")
            else: 
                filename = "data"
            df = pd.read_csv(f"{path}{filename}.csv", usecols=self.config["target_cols"])

            # 발송 여부 확인
            while True:
                check = str(input(f"{len(df)}명에게 정말로 발송하시겠습니까? (Y/N)")).lower()
                if check == "n":
                    raise UserCancelException
                elif check == "y":
                    break
                else:
                    continue
            
            for index, row in df.iterrows():
                if index >= 500: 
                    logging.info(f"mail send limit exceeded: {index}")
                    break
                user_email = row["이메일"]
                user_name = row["업체명"]

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
                    pdf_part.add_header("Content-Disposition", "attachment", filename=f"{self.config["pdf_filename"]}.pdf")
                    msg.attach(pdf_part)

                    self.server.send_message(msg)
                    logging.info(f"send email to {user_email}({user_name})")

                except Exception as e:
                    logging.error(f"failed send email: {e}")
                    continue

                finally:
                    logging.info(f"complete! {index}")
                    # time.sleep(1.5)

        except UserCancelException:
            logging.info("user canceled.")

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
        
    except Exception as e:
        logging.critical(f"fatal error detected: {e}")
        sys.exit()

    finally:
        logging.info("program close.")
        app.close()