from models import db, AuditLog
from flask import request, session


def log_action(action, user_id=None, user_email=None, role=None):
    """
    Log a system action to the audit log.
    Falls back to session values when not explicitly provided,
    allowing pre-auth security events (failed login, lockout, etc.)
    to be recorded with None user_id.
    """
    try:

        log = AuditLog(
            user_id=user_id if user_id is not None else session.get("user_id"),
            user_email=user_email if user_email is not None else session.get("user_email"),
            role=role if role is not None else session.get("role"),
            action=action,
            ip_address=request.remote_addr
        )

        db.session.add(log)
        db.session.commit()

    except Exception as e:
        print("Audit log failed:", e)
