import os
from datetime import timedelta
from dotenv import load_dotenv

load_dotenv()  # load .env file automatically


class Config:

    # =====================================================
    # Flask Security
    # =====================================================

    SECRET_KEY = os.environ.get("SECRET_KEY")

    if not SECRET_KEY:
        import secrets as _s
        SECRET_KEY = _s.token_hex(32)
        import warnings
        warnings.warn(
            "SECRET_KEY not set in environment. "
            "A random key was generated — sessions will be lost on restart. "
            "Set SECRET_KEY in your .env file for production.",
            stacklevel=2
        )

    # Session cookie hardening
    SESSION_COOKIE_HTTPONLY = True        # JS cannot access cookie
    SESSION_COOKIE_SAMESITE = "Lax"      # CSRF baseline
    SESSION_COOKIE_SECURE = os.environ.get("HTTPS", "false").lower() == "true"

    PERMANENT_SESSION_LIFETIME = timedelta(minutes=30)   # auto-logout

    WTF_CSRF_ENABLED = True
    WTF_CSRF_TIME_LIMIT = 3600           # 1 hour

    TEMPLATES_AUTO_RELOAD = True         # always pick up template changes

    # =====================================================
    # Account Lockout
    # =====================================================
    MAX_LOGIN_ATTEMPTS = 5               # fails before lockout
    LOCKOUT_MINUTES = 15                 # lockout duration

    # =====================================================
    # Allowed upload extensions
    # =====================================================
    ALLOWED_EXTENSIONS = {
        "pdf", "doc", "docx", "jpg", "jpeg", "png", "xlsx", "xls"
    }


    # =====================================================
    # PostgreSQL Database Configuration
    # =====================================================

    DB_USER     = os.environ.get("DB_USER", "postgres")
    DB_PASSWORD  = os.environ.get("DB_PASSWORD")
    DB_HOST      = os.environ.get("DB_HOST", "localhost")
    # DB_PORT      = os.environ.get("DB_PORT", "5432")
    DB_NAME      = os.environ.get("DB_NAME", "faculty_mis_db")


    SQLALCHEMY_DATABASE_URI = (
        f"postgresql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}/{DB_NAME}?sslmode=require"
    )

    SQLALCHEMY_TRACK_MODIFICATIONS = False


    # =====================================================
    # Gmail SMTP Configuration (FIXED & PRODUCTION SAFE)
    # =====================================================

    MAIL_SERVER = "smtp.gmail.com"

    MAIL_PORT = 587

    MAIL_USE_TLS = True

    MAIL_USE_SSL = False


    MAIL_USERNAME       = os.environ.get("MAIL_USERNAME")
    MAIL_PASSWORD       = os.environ.get("MAIL_PASSWORD")
    MAIL_DEFAULT_SENDER = os.environ.get("MAIL_DEFAULT_SENDER")


    # IMPORTANT FIX FOR CONNECTION ERROR
    MAIL_SUPPRESS_SEND = False

    MAIL_DEBUG = os.environ.get("FLASK_DEBUG", "false").lower() == "true"

    MAIL_TIMEOUT = 30


    # =====================================================
    # Fast2SMS Configuration
    # =====================================================

    FAST2SMS_API_KEY = os.environ.get("FAST2SMS_API_KEY")


    # =====================================================
    # OTP Settings
    # =====================================================

    OTP_EXPIRY_MINUTES = 5

    OTP_LENGTH = 6


    # =====================================================
    # File Upload Settings
    # =====================================================

    UPLOAD_FOLDER = "uploads"

    EXPORT_FOLDER = "exports"

    MAX_CONTENT_LENGTH = 16 * 1024 * 1024   # 16 MB limit
