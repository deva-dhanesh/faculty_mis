import pandas as pd
from datetime import datetime
from models import db, User, FacultyProfile
from security_utils import hash_password


# =========================================================
# CSV INJECTION SANITIZATION
# =========================================================

_CSV_INJECTION_PREFIXES = ('=', '+', '-', '@', '\t', '\r')

def _safe(value):
    """Prefix dangerous values to prevent CSV formula injection."""
    if value is None:
        return value
    s = str(value)
    if s.startswith(_CSV_INJECTION_PREFIXES):
        return "'" + s
    return s


# =========================================================
# BULK CREATE FACULTY USERS (AUTH TABLE)
# =========================================================
def upload_faculty_csv(file):

    df = pd.read_csv(file)

    added = 0

    for _, row in df.iterrows():

        employee_id = str(row["employee_id"])
        email = row["email"]
        phone = str(row["phone"])
        password = row["password"]

        existing = User.query.filter(
            (User.email == email) | (User.employee_id == employee_id)
        ).first()

        if existing:
            continue

        user = User(
            employee_id=employee_id,
            email=email,
            phone=phone,
            password_hash=hash_password(password),
            role="faculty"
        )

        db.session.add(user)
        added += 1

    db.session.commit()

    return added


# =========================================================
# EXPORT FACULTY USERS CSV
# =========================================================
def export_faculty_csv(filepath):

    users = User.query.filter_by(role="faculty").all()

    data = []

    for user in users:
        data.append({
            "employee_id": _safe(user.employee_id),
            "email": _safe(user.email),
            "phone": _safe(user.phone)
        })

    df = pd.DataFrame(data)
    df.to_csv(filepath, index=False)


# =========================================================
# BULK UPLOAD FACULTY PROFILES
# =========================================================
def upload_faculty_profiles_csv(file):

    df = pd.read_csv(file)

    added = 0

    for _, row in df.iterrows():

        employee_id = str(row["employee_id"])

        user = User.query.filter_by(employee_id=employee_id).first()

        # skip if user not found
        if not user:
            continue

        profile = FacultyProfile.query.filter_by(
            user_id=user.id
        ).first()

        if not profile:
            profile = FacultyProfile(user_id=user.id)

        # Assign fields safely
        profile.employee_id = employee_id
        profile.full_name = row.get("full_name")
        profile.pan = row.get("pan")
        profile.designation = row.get("designation")

        # Date parsing safely
        doj = row.get("date_of_joining")
        dob = row.get("date_of_birth")

        if pd.notna(doj):
            profile.date_of_joining = datetime.strptime(
                str(doj), "%Y-%m-%d"
            ).date()

        if pd.notna(dob):
            profile.date_of_birth = datetime.strptime(
                str(dob), "%Y-%m-%d"
            ).date()

        profile.appointment_nature = row.get("appointment_nature")
        profile.qualification = row.get("qualification")
        profile.department = row.get("department")

        exp = row.get("experience_years")
        if pd.notna(exp):
            profile.experience_years = int(exp)

        profile.mobile = str(row.get("mobile"))
        profile.email_personal = row.get("email_personal")
        profile.email_university = row.get("email_university")

        profile.specialization = row.get("specialization")

        profile.appointment_letter_url = row.get(
            "appointment_letter_url"
        )

        db.session.add(profile)

        added += 1

    db.session.commit()

    return added


# =========================================================
# EXPORT FACULTY PROFILES CSV
# =========================================================
def export_faculty_profiles_csv(filepath):

    profiles = FacultyProfile.query.all()

    data = []

    for p in profiles:

        data.append({

            "employee_id": _safe(p.employee_id),
            "full_name": _safe(p.full_name),
            "pan": _safe(p.pan),
            "designation": _safe(p.designation),

            "date_of_joining": p.date_of_joining,
            "date_of_birth": p.date_of_birth,

            "appointment_nature": _safe(p.appointment_nature),

            "qualification": _safe(p.qualification),
            "department": _safe(p.department),

            "experience_years": p.experience_years,

            "mobile": _safe(p.mobile),
            "email_personal": _safe(p.email_personal),
            "email_university": _safe(p.email_university),

            "specialization": _safe(p.specialization),

            "appointment_letter_url": _safe(p.appointment_letter_url)
        })

    df = pd.DataFrame(data)

    df.to_csv(filepath, index=False)
