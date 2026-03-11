import secrets
import bcrypt
import requests

from datetime import datetime, timedelta

from flask_mail import Message
from flask import current_app

from models import db, User


# =====================================================
# Generate OTP  (cryptographically secure)
# =====================================================

def generate_otp():

    length = current_app.config.get("OTP_LENGTH", 6)

    return "".join(
        secrets.choice("0123456789")
        for _ in range(length)
    )


# =====================================================
# Store OTP (Hashed)
# =====================================================

def store_otp(user: User, otp: str, otp_type="login"):

    otp_hash = bcrypt.hashpw(
        otp.encode(),
        bcrypt.gensalt()
    ).decode()

    expiry_minutes = current_app.config.get(
        "OTP_EXPIRY_MINUTES", 5
    )

    expiry = datetime.utcnow() + timedelta(
        minutes=expiry_minutes
    )

    if otp_type == "login":

        user.otp_hash = otp_hash
        user.otp_expiry = expiry

    else:

        user.reset_otp_hash = otp_hash
        user.reset_otp_expiry = expiry

    db.session.commit()


# =====================================================
# Verify OTP
# =====================================================

def verify_otp(user: User, entered_otp: str, otp_type="login"):

    if otp_type == "login":

        otp_hash = user.otp_hash
        expiry = user.otp_expiry

    else:

        otp_hash = user.reset_otp_hash
        expiry = user.reset_otp_expiry

    if not otp_hash or not expiry:
        return False

    if datetime.utcnow() > expiry:
        return False

    return bcrypt.checkpw(
        entered_otp.encode(),
        otp_hash.encode()
    )


# =====================================================
# Send Email OTP (FIXED)
# =====================================================

def send_email_otp(user_email: str, otp: str):

    try:

        # Import inside function to avoid circular import
        from app import mail

        msg = Message(

            subject="Faculty MIS OTP Verification",

            sender=current_app.config["MAIL_DEFAULT_SENDER"],

            recipients=[user_email]

        )

        msg.body = f"""
Dear User,

Your Faculty MIS OTP is: {otp}

This OTP is valid for {current_app.config.get("OTP_EXPIRY_MINUTES",5)} minutes.

Do NOT share this OTP.

Regards,
Faculty MIS System
"""

        mail.send(msg)

        print(f"✅ Email OTP sent to: {user_email}")

    except Exception as e:

        print("❌ Email sending failed:", str(e))


# =====================================================
# Send SMS OTP (Fast2SMS FIXED)
# =====================================================

def send_sms_otp(phone_number: str, otp: str):

    try:

        api_key = current_app.config["FAST2SMS_API_KEY"]

        url = "https://www.fast2sms.com/dev/bulkV2"

        message = f"Your Faculty MIS OTP is {otp}. Valid for 5 minutes."

        params = {

            "authorization": api_key,

            "route": "q",

            "message": message,

            "language": "english",

            "numbers": phone_number

        }

        headers = {

            "cache-control": "no-cache"

        }

        response = requests.get(
            url,
            params=params,
            headers=headers
        )

        print("📱 SMS response:", response.json())

    except Exception as e:

        print("❌ SMS sending failed:", str(e))


# =====================================================
# Send BOTH Email and SMS
# =====================================================

def send_otp(user: User, otp: str):

    send_email_otp(user.email, otp)

    send_sms_otp(user.phone, otp)
