import os

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))


class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')
    DATABASE = os.path.join(BASE_DIR, 'instance', 'modeltiering.db')
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'instance', 'uploads')
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024
