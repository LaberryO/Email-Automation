from enum import Enum, auto

class ScreenType(Enum):
    MAIN = [
        "--- System Console ---",
        "파일 업로드",
        "이메일 발송"
    ]
    LOAD_DATA = [
        "--- Load Data ---",
        "파일 경로: "
    ]
    SEND_EMAIL = [
        "--- Send Email ---"
    ]
    EXIT_PROGRAM = [
        "--- Exit Program ---"
    ]

