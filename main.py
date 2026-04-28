import pandas as pd
import smtplib, sys, json, datetime, time, logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

# =======
# setting
# =======

now = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(f"logs/log_{now}.txt", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

FILE_NAME = ["data.json", "custom_data.csv"] # index 0: 사용자 정보, index 1: 고객 정보
TARGET_COLS = ["업체명", "이메일"]

EMAIL_CONTENT = """
<html>
    <body>
        <img src="cid:image">
    </body>
</html>
"""

success_list = []
fail_list = []

# =======
# load data
# =======

logging.info("program start")

try:
    with open(f"localuser/{FILE_NAME[0]}", "r", encoding="utf-8") as file:
        data = json.load(file)
        logging.info("data load complete")

except FileNotFoundError:
    logging.info("data file not found")

except json.JSONDecodeError:
    logging.error("json decode failed")

except PermissionError:
    logging.error("no permission")

# =======
# smtp connection
# =======

try:
    logging.info("try connect smtp server..")
    sv = smtplib.SMTP_SSL(data["SMTP_ADDRESS"], data["SMTP_PORT"])
    sv.login(data["EMAIL"], data["APP_PASSWORD"])
    logging.info("stmp connected")

except smtplib.SMTPAuthenticationError:
    logging.error("incorrect id and password")
    sys.exit()

except Exception as e:
    logging.error(f"smtp connection error: {e}")
    sys.exit()

# =======
# Email Logic
# =======

try:
    with open("localuser/image.png", "rb") as f:
        cached_image_data = f.read()
    
    with open("localuser/catalog.pdf", "rb") as f:
        cached_pdf_data = f.read()

    df = pd.read_csv(f"localuser/{FILE_NAME[1]}", usecols=TARGET_COLS)

    for index, row in df.iterrows():
        user_email = row["이메일"]
        user_name = row["업체명"]

        # send email
        try:
            msg = MIMEMultipart()
            msg["Subject"] = data["EMAIL_SUBJECT"]
            msg["From"] = data["EMAIL"]
            msg["To"] = user_email

            msg.attach(MIMEText(EMAIL_CONTENT, "html"))

            # 본문에 이미지 넣기
            image_part = MIMEImage(cached_image_data)
            image_part.add_header("Content-ID", "<image>")
            msg.attach(image_part)

            # PDF 파일 첨부
            pdf_part = MIMEApplication(cached_pdf_data, _subtype="pdf")
            pdf_part.add_header("Content-Disposition", "attachment", filename=data["PDF_FILENAME"])
            msg.attach(pdf_part)

            sv.send_message(msg)
            logging.info(f"send email to {user_email}({user_name})")
            success_list.append(f"{user_email}")
        
        except Exception as e:
            logging.error(f"failed send email: {e}")
            fail_list.append(f"{user_email} reason: {e}")
            continue

        finally:
            logging.info("wait a second..")
            time.sleep(1.5)

except FileNotFoundError:
    logging.error(f"{FILE_NAME[1]} not found")

except Exception as e:
    logging.error(f"email logic error: {e}")

finally:
    # sever disconnect
    try:
        sv.quit()
        logging.info("server closed. program ended.")
    except Exception as e:
        logging.error("abnormally shutdown detected. program ended.")