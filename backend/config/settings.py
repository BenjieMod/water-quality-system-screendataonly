import os
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
load_dotenv(os.path.join(BASE_DIR, '.env'))

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'dev-secret-key-change-in-production'
    SQLALCHEMY_DATABASE_URI = 'sqlite:///water_quality.db'
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    APP_AUTH_REQUIRED = os.environ.get('APP_AUTH_REQUIRED', 'true').lower() == 'true'
    APP_DEFAULT_ADMIN_USERNAME = os.environ.get('APP_DEFAULT_ADMIN_USERNAME', 'admin')
    APP_DEFAULT_ADMIN_PASSWORD = os.environ.get('APP_DEFAULT_ADMIN_PASSWORD', 'admin123')
    APP_SESSION_COOKIE_SECURE = os.environ.get('APP_SESSION_COOKIE_SECURE', 'false').lower() == 'true'
    MONITORING_LOGIN_URL = os.environ.get('MONITORING_LOGIN_URL', 'http://192.168.1.152:8082/production/pages/login.jsp')
    MONITORING_USERNAME = os.environ.get('MONITORING_USERNAME', '')
    MONITORING_PASSWORD = os.environ.get('MONITORING_PASSWORD', '')
    MONITORING_HEADLESS = os.environ.get('MONITORING_HEADLESS', 'true').lower() == 'true'
    MONITORING_ENTRY_DELAY_MINUTES = int(os.environ.get('MONITORING_ENTRY_DELAY_MINUTES', '12'))
