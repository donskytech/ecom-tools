import os


class Config:
    SECRET_KEY = os.getenv("SECRET_KEY", "dev-secret-key")
    MAX_CONTENT_LENGTH = 10 * 1024 * 1024  # 10MB uploads
