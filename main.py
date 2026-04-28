import pandas as pd
import smtplib, sys, json, datetime, time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

# =======
# setting
# =======

FILE_NAME = ["data.json", "test.csv"] # index 0: 사용자 정보, index 1: 고객 정보
TARGET_COLS = ["업체명", "이메일"]

EMAIL_CONTENT = """
<html>
    <body>
        <img src="cid:image">
    </body>
</html>
"""

STATUS = ["[INFO]", "[SUCCESS]", "[ERROR]"]

success_list = []
fail_list = []

# =======
# load data
# =======

def exception(text, status):
    print(f"{STATUS[status]} {text}")

try:
    with open(f"localuser/{FILE_NAME[0]}", "r", encoding="utf-8") as file:
        data = json.load(file)
        exception("data load complete",1)

except FileNotFoundError:
    exception("data file not found",2)

except json.JSONDecodeError:
    exception("json decode failed",2)

except PermissionError:
    exception("no permission",2)

# =======
# smtp connection
# =======

try:
    exception("try connect smtp server..",0)
    sv = smtplib.SMTP_SSL(data["SMTP_ADDRESS"], data["SMTP_PORT"])
    sv.login(data["EMAIL"], data["APP_PASSWORD"])
    exception("stmp connected",1)

except smtplib.SMTPAuthenticationError:
    exception("incorrect id and password",2)
    sys.exit()

except Exception as e:
    exception(f"smtp connection error: {e}", 2)
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
            exception(f"send email to {user_email}({user_name})",1)
            success_list.append(f"{user_email}")
        
        except Exception as e:
            exception(f"failed send email: {e}",2)
            fail_list.append(f"{user_email} reason: {e}")
            continue

        finally:
            exception("wait a second..",0)
            time.sleep(1.5)

except FileNotFoundError:
    exception(f"{FILE_NAME[1]} not found",2)

except Exception as e:
    exception(f"email logic error: {e}",2)

finally:
    # sever disconnect
    try:
        sv.quit()
        exception("server closed",0)
    except:
        pass

    # log
    try:
        # 현재 시간을 'YYYY-MM-DD_HH-MM-SS' 형태로 포맷팅
        now = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        log_filename = f"log_{now}.txt"
        
        # 이전 질문에서 다뤘던 'with open'을 활용해 파일 작성 ('w' 모드)
        with open(f"logs/{log_filename}", "w", encoding="utf-8") as log_file:
            log_file.write(f"{now}\n")
            log_file.write("\n")
            
            # 성공 내역 기록
            log_file.write(f"[SUCCESS] {len(success_list)}\n")
            for email in success_list:
                log_file.write(f" - {email}\n")
                
            log_file.write("\n")
            
            # 실패 내역 기록
            log_file.write(f"[ERROR] {len(fail_list)}\n")
            for fail_info in fail_list:
                log_file.write(f" - {fail_info}\n")
        
        exception("program ended. log file saved.",0)

    except Exception as e:
        exception("program ended. but log file save failed",2)