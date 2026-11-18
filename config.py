import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'chave-interna-conferencia-fiscal-2026')
    DATABASE_PATH = os.path.join(BASE_DIR, 'fiscal_audit.db')
    UPLOAD_RAW_FOLDER = os.path.join(BASE_DIR, 'uploads', 'raw')
    REPORTS_FOLDER = os.path.join(BASE_DIR, 'uploads', 'reports')
    ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # Limite de 50 MB por envio

    @staticmethod
    def init_app():
        os.makedirs(Config.UPLOAD_RAW_FOLDER, exist_ok=True)
        os.makedirs(Config.REPORTS_FOLDER, exist_ok=True)
