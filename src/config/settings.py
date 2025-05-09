import os
from dotenv import load_dotenv

# .env에서 모든 환경변수 불러오기
load_dotenv()

# Dropbox 관련 환경변수
DROPBOX_APP_KEY = os.getenv("DROPBOX_APP_KEY", "")
DROPBOX_APP_SECRET = os.getenv("DROPBOX_APP_SECRET", "")
DROPBOX_ACCESS_TOKEN = os.getenv("DROPBOX_ACCESS_TOKEN", "")
DROPBOX_REFRESH_TOKEN = os.getenv("DROPBOX_REFRESH_TOKEN", "")
DROPBOX_SHARED_FOLDER_ID = os.getenv("DROPBOX_SHARED_FOLDER_ID", "")
DROPBOX_SHARED_FOLDER_NAME = os.getenv("DROPBOX_SHARED_FOLDER_NAME", "")
LOCAL_BID_FOLDER = os.getenv("LOCAL_BID_FOLDER", "")

# ChatGPT 관련 환경변수
GPT_API_KEY = os.getenv("CHATGPT_API_KEY", "")
GPT_MODEL = os.getenv("CHATGPT_MODEL", "gpt-4.1-mini") 