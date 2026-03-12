import io
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from config import Config
from models import (
    db,
    User,
    FacultyProfile,
    FacultyDocument,
    FacultyPublication,
    FacultyProject,
    FacultyResignation,
    FacultyDegree,
    AuditLog,
    CourseAttended,
    CourseOffered,
    FacultyAward,
    FacultyBookChapter,
    FacultyGuestLecture,
    FacultyPatent,
    FacultyFellowship,
    ConferenceParticipated,
    ConferenceOrganised,
    FDPParticipated,
    FDPOrganised
)

from flask_login import LoginManager
from flask_mail import Mail
from flask_wtf.csrf import CSRFProtect, CSRFError
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

from security_utils import verify_password, hash_password
from otp_utils import generate_otp, store_otp, verify_otp, send_otp
from audit_utils import log_action
from report_utils import generate_admin_report, generate_personal_report
from bulk_utils import (
    build_publications_template, build_projects_template,
    parse_publications, parse_projects,
    build_faculty_users_template, build_faculty_profiles_template,
    parse_faculty_users, parse_faculty_profiles,
    build_courses_attended_template, build_courses_offered_template,
    parse_courses_attended, parse_courses_offered,
    build_awards_template, parse_awards,
    build_book_chapters_template, parse_book_chapters,
    build_guest_lectures_template, parse_guest_lectures,
    build_patents_template, parse_patents,
    build_fellowships_template, parse_fellowships,
    build_conferences_participated_template, parse_conferences_participated,
    build_conferences_organised_template, parse_conferences_organised,
    build_fdp_participated_template, parse_fdp_participated,
    build_fdp_organised_template, parse_fdp_organised
)
from csv_utils import (
    export_faculty_csv,
    export_faculty_profiles_csv
)

from sqlalchemy import func
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
import os
import uuid


# =========================================================
# APP INITIALIZATION
# =========================================================

app = Flask(__name__)

# Load config
app.config.from_object(Config)

# Make sessions permanent so PERMANENT_SESSION_LIFETIME applies
@app.before_request
def make_session_permanent():
    session.permanent = True


# =========================================================
# PER-REQUEST ACTIVE ACCOUNT VERIFICATION
# =========================================================

@app.before_request
def verify_active_account():
    """
    On every request, if a user is logged in, verify their account
    is still active in the DB. Handles mid-session deactivation.
    """
    user_id = session.get("user_id")
    if user_id:
        try:
            user = User.query.get(user_id)
            if not user or not user.is_active_account:
                session.clear()
                flash("Your account has been deactivated. Please contact admin.", "error")
                return redirect(url_for("login"))
        except Exception:
            # DB error — rollback, clear session, redirect to login
            db.session.rollback()
            session.clear()
            return redirect(url_for("login"))


@app.teardown_request
def teardown_request_handler(exception):
    """Always rollback on any unhandled exception so the connection
    pool never returns a broken 'transaction aborted' connection."""
    if exception is not None:
        db.session.rollback()

# Initialize extensions
db.init_app(app)
mail = Mail(app)
csrf = CSRFProtect(app)

limiter = Limiter(
    key_func=get_remote_address,
    app=app,
    default_limits=[],
    storage_uri="memory://"
)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"


# =========================================================
# SECURITY HEADERS  (applied to every response)
# =========================================================

@app.after_request
def set_security_headers(response):
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-XSS-Protection"] = "1; mode=block"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Permissions-Policy"] = "geolocation=(), microphone=(), camera=()"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net https://cdnjs.cloudflare.com; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com https://cdnjs.cloudflare.com; "
        "font-src 'self' https://fonts.gstatic.com https://cdnjs.cloudflare.com; "
        "img-src 'self' data:;"
    )
    if app.config.get("SESSION_COOKIE_SECURE"):
        response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    return response


# =========================================================
# CSRF ERROR HANDLER
# =========================================================

@app.errorhandler(CSRFError)
def handle_csrf_error(e):
    flash("Session expired or invalid request. Please try again.", "error")
    return redirect(url_for("login")), 400


# =========================================================
# HELPER — allowed file extension check
# =========================================================

ALLOWED_EXTENSIONS = app.config.get(
    "ALLOWED_EXTENSIONS",
    {"pdf", "doc", "docx", "jpg", "jpeg", "png"}
)

def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
    )


# =========================================================
# FOLDER SETUP
# =========================================================

UPLOAD_FOLDER = app.config.get("UPLOAD_FOLDER", "uploads")
EXPORT_FOLDER = app.config.get("EXPORT_FOLDER", "exports")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)


# =========================================================
# LOGIN MANAGER USER LOADER
# =========================================================

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():
    return redirect(url_for("login"))


# =========================================================
# LOGIN
# templates/auth/login.html
# =========================================================

@app.route("/login", methods=["GET", "POST"])
@limiter.limit("10 per minute")
def login():

    # If already logged in
    if session.get("user_id"):
        if session.get("role") == "admin":
            return redirect(url_for("admin_dashboard"))
        else:
            return redirect(url_for("faculty_dashboard"))

    if request.method == "POST":

        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")

        user = User.query.filter_by(email=email).first()

        # Generic message — prevents user enumeration
        _fail_msg = "Invalid email or password."

        if not user or not user.is_active_account:
            flash(_fail_msg, "error")
            return redirect(url_for("login"))

        # Check lockout
        if user.locked_until and datetime.utcnow() < user.locked_until:
            remaining = int((user.locked_until - datetime.utcnow()).total_seconds() // 60) + 1
            flash(f"Account locked. Try again in {remaining} minute(s).", "error")
            log_action(
                f"Login blocked — account locked: {email}",
                user_email=email
            )
            return redirect(url_for("login"))

        if not verify_password(password, user.password_hash):

            user.failed_attempts = (user.failed_attempts or 0) + 1
            max_attempts = app.config.get("MAX_LOGIN_ATTEMPTS", 5)

            if user.failed_attempts >= max_attempts:
                lockout_mins = app.config.get("LOCKOUT_MINUTES", 15)
                user.locked_until = datetime.utcnow() + timedelta(minutes=lockout_mins)
                db.session.commit()
                log_action(
                    f"Account locked after {max_attempts} failed attempts: {email}",
                    user_email=email
                )
                flash(f"Too many failed attempts. Account locked for {lockout_mins} minutes.", "error")
            else:
                db.session.commit()
                flash(_fail_msg, "error")

            return redirect(url_for("login"))

        # Successful password — reset counters
        user.failed_attempts = 0
        user.locked_until = None
        db.session.commit()

        # ── FACULTY: first-time login → OTP then force-reset password ──
        if user.role == "faculty" and user.must_reset_password:
            otp = generate_otp()
            store_otp(user, otp)
            session["otp_user_id"] = user.id
            session["otp_next"] = "force_reset"
            send_otp(user, otp)
            flash("Welcome! An OTP has been sent to your email. After verification you will set your own password.", "success")
            return redirect(url_for("otp_verification"))

        # ── FACULTY: normal login — no OTP required ──────────────────
        if user.role == "faculty":
            session["user_id"]    = user.id
            session["role"]       = user.role
            session["user_email"] = user.email
            log_action("Faculty logged in")
            flash("Login successful.", "success")
            return redirect(url_for("faculty_dashboard"))

        # ── ADMIN: always OTP ─────────────────────────────────────────
        otp = generate_otp()
        store_otp(user, otp)
        session["otp_user_id"] = user.id
        send_otp(user, otp)
        flash("OTP sent to your email.", "success")
        return redirect(url_for("otp_verification"))

    return render_template("auth/login.html")


# =========================================================
# OTP VERIFY
# templates/auth/otp.html
# =========================================================

@app.route("/otp", methods=["GET", "POST"])
@limiter.limit("10 per minute")
def otp_verification():

    if request.method == "POST":

        entered_otp = request.form.get("otp", "").strip()

        user_id = session.get("otp_user_id")

        if not user_id:
            flash("Session expired. Please log in again.", "error")
            return redirect(url_for("login"))

        user = User.query.get(user_id)

        if not user:
            session.clear()
            flash("Invalid session.", "error")
            return redirect(url_for("login"))

        if verify_otp(user, entered_otp):

            # Clear OTP fields after successful use
            user.otp_hash = None
            user.otp_expiry = None
            user.otp_attempts = 0
            db.session.commit()

            session.pop("otp_user_id", None)

            # Regenerate session to prevent session fixation
            uid = user.id
            role = user.role
            email = user.email
            next_page = session.pop("otp_next", None)   # read BEFORE clear
            session.clear()

            session["user_id"]    = uid
            session["role"]       = role
            session["user_email"] = email

            log_action("User verified OTP")

            # First-time faculty → force password reset
            if next_page == "force_reset":
                session["force_reset_user_id"] = uid
                return redirect(url_for("force_reset_password"))

            flash("Login successful.", "success")

            if role == "admin":
                return redirect(url_for("admin_dashboard"))
            else:
                return redirect(url_for("faculty_dashboard"))

        # Wrong OTP — track attempts
        user.otp_attempts = (user.otp_attempts or 0) + 1
        max_otp = app.config.get("MAX_LOGIN_ATTEMPTS", 5)

        if user.otp_attempts >= max_otp:
            # Invalidate OTP entirely — user must log in again
            user.otp_hash = None
            user.otp_expiry = None
            user.otp_attempts = 0
            db.session.commit()
            session.clear()
            log_action("OTP invalidated after too many failed attempts")
            flash("Too many incorrect OTP attempts. Please log in again.", "error")
            return redirect(url_for("login"))

        db.session.commit()
        log_action("Failed OTP attempt")
        flash("Invalid or expired OTP.", "error")

    return render_template("auth/otp.html")


# =========================================================
# FORCE RESET PASSWORD (Faculty first-time login)
# templates/auth/force_reset_password.html
# =========================================================

@app.route("/force_reset_password", methods=["GET", "POST"])
def force_reset_password():

    user_id = session.get("force_reset_user_id")

    if not user_id or not session.get("user_id"):
        flash("Session expired. Please log in again.", "error")
        return redirect(url_for("login"))

    user = User.query.get(user_id)

    if not user:
        session.clear()
        return redirect(url_for("login"))

    if request.method == "POST":

        new_password     = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")

        if len(new_password) < 8:
            flash("Password must be at least 8 characters.", "error")
            return redirect(url_for("force_reset_password"))

        if new_password != confirm_password:
            flash("Passwords do not match.", "error")
            return redirect(url_for("force_reset_password"))

        user.password_hash      = hash_password(new_password)
        user.must_reset_password = False
        db.session.commit()

        session.pop("force_reset_user_id", None)
        log_action("Faculty set their password on first login")
        flash("Password set successfully. Welcome!", "success")
        return redirect(url_for("faculty_dashboard"))

    return render_template("auth/force_reset_password.html", user=user)


# =========================================================
# ADMIN DASHBOARD
# templates/admin/admin_dashboard.html
# =========================================================

@app.route("/admin_dashboard")
def admin_dashboard():

    if session.get("role") != "admin":
        flash("Access denied", "error")
        return redirect(url_for("login"))

    log_action("Admin opened dashboard")

    faculty_count    = User.query.filter_by(role="faculty").count()
    profile_count    = FacultyProfile.query.count()
    publication_count = FacultyPublication.query.count()
    project_count    = FacultyProject.query.count()
    document_count   = FacultyDocument.query.count()
    book_chapter_count = FacultyBookChapter.query.count()
    guest_lecture_count = FacultyGuestLecture.query.count()
    patent_count = FacultyPatent.query.count()
    fellowship_count = FacultyFellowship.query.count()
    conf_participated_count = ConferenceParticipated.query.count()
    conf_organised_count = ConferenceOrganised.query.count()
    fdp_participated_count = FDPParticipated.query.count()
    fdp_organised_count = FDPOrganised.query.count()

    return render_template(
        "admin/admin_dashboard.html",
        faculty_count=faculty_count,
        profile_count=profile_count,
        publication_count=publication_count,
        project_count=project_count,
        document_count=document_count,
        book_chapter_count=book_chapter_count,
        guest_lecture_count=guest_lecture_count,
        patent_count=patent_count,
        fellowship_count=fellowship_count,
        conf_participated_count=conf_participated_count,
        conf_organised_count=conf_organised_count,
        fdp_participated_count=fdp_participated_count,
        fdp_organised_count=fdp_organised_count,
    )


# =========================================================
# FACULTY DASHBOARD
# templates/faculty/faculty_dashboard.html
# =========================================================

@app.route("/faculty_dashboard")
def faculty_dashboard():

    if session.get("role") != "faculty":
        flash("Access denied", "error")
        return redirect(url_for("login"))

    log_action("Faculty opened dashboard")

    return render_template("faculty/faculty_dashboard.html")


# =========================================================
# GENERATE FACULTY REPORT
# templates/faculty/generate_report.html
# =========================================================

@app.route("/generate_report")
def generate_report_page():
    """Display report generation selection page"""
    
    if session.get("role") != "faculty":
        flash("Access denied", "error")
        return redirect(url_for("login"))
    
    log_action("Faculty accessed report generation page")
    
    return render_template("faculty/generate_report.html")


@app.route("/generate_faculty_report", methods=["POST"])
def generate_faculty_report():
    """Generate and display faculty report"""
    
    if session.get("role") != "faculty":
        flash("Access denied", "error")
        return redirect(url_for("login"))
    
    # Get selected features from form
    selected_features = request.form.getlist("features")
    
    if not selected_features:
        flash("Please select at least one feature.", "error")
        return redirect(url_for("generate_report_page"))
    
    try:
        from report_generation import (
            generate_charts, generate_interpretation, 
            compile_summary, generate_detailed_stats
        )
        
        user_id = session["user_id"]
        
        # Generate all report components
        charts = generate_charts(user_id, selected_features)
        interpretation = generate_interpretation(user_id, selected_features)
        summary = compile_summary(user_id, selected_features)
        detailed_stats = generate_detailed_stats(user_id, selected_features)
        
        log_action(f"Faculty generated report with features: {', '.join(selected_features)}")
        
        return render_template(
            "faculty/view_report.html",
            charts=charts,
            interpretation=interpretation,
            summary=summary,
            detailed_stats=detailed_stats,
            current_date=datetime.now().strftime("%d %B %Y")
        )
    
    except Exception as e:
        flash(f"Error generating report: {str(e)}", "error")
        return redirect(url_for("generate_report_page"))


# =========================================================
# VIEW FACULTY PROFILES
# templates/admin/view_faculty_profiles.html
# =========================================================

@app.route("/view_faculty_profiles")
def view_faculty_profiles():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    q        = request.args.get("q", "").strip()
    dept     = request.args.get("dept", "").strip()
    desig    = request.args.get("desig", "").strip()
    exp_min  = request.args.get("exp_min", "").strip()

    qry = FacultyProfile.query
    if q:
        qry = qry.filter(db.or_(
            FacultyProfile.full_name.ilike(f"%{q}%"),
            FacultyProfile.employee_id.ilike(f"%{q}%"),
            FacultyProfile.specialization.ilike(f"%{q}%"),
        ))
    if dept:
        qry = qry.filter(FacultyProfile.department == dept)
    if desig:
        qry = qry.filter(FacultyProfile.designation == desig)
    if exp_min and exp_min.isdigit():
        qry = qry.filter(FacultyProfile.experience_years >= int(exp_min))

    profiles     = qry.order_by(FacultyProfile.full_name).all()
    departments  = sorted({p.department  for p in FacultyProfile.query.all() if p.department})
    designations = sorted({p.designation for p in FacultyProfile.query.all() if p.designation})

    log_action("Admin viewed faculty profiles")

    return render_template(
        "admin/view_faculty_profiles.html",
        profiles=profiles,
        departments=departments,
        designations=designations,
        q=q, dept=dept, desig=desig, exp_min=exp_min,
    )


# =========================================================
# VIEW DOCUMENTS
# templates/admin/view_documents.html
# =========================================================

@app.route("/view_documents")
def view_documents():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    q        = request.args.get("q", "").strip()
    doc_type = request.args.get("doc_type", "").strip()

    qry = FacultyDocument.query
    if q:
        qry = qry.filter(db.or_(
            FacultyDocument.file_name.ilike(f"%{q}%"),
            FacultyDocument.document_type.ilike(f"%{q}%"),
        ))
    if doc_type:
        qry = qry.filter(FacultyDocument.document_type == doc_type)

    documents = qry.order_by(FacultyDocument.uploaded_at.desc()).all()
    doc_types = sorted({d.document_type for d in FacultyDocument.query.all() if d.document_type})

    log_action("Admin viewed documents")

    return render_template(
        "admin/view_documents.html",
        documents=documents,
        doc_types=doc_types,
        q=q, doc_type=doc_type,
    )


# =========================================================
# VIEW PUBLICATIONS
# templates/admin/view_publications.html
# =========================================================

@app.route("/view_publications")
def view_publications():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    q       = request.args.get("q", "").strip()
    year    = request.args.get("year", "").strip()
    indexing = request.args.get("indexing", "").strip()

    qry = FacultyPublication.query
    if q:
        qry = qry.filter(db.or_(
            FacultyPublication.title.ilike(f"%{q}%"),
            FacultyPublication.author_position.ilike(f"%{q}%"),
            FacultyPublication.journal.ilike(f"%{q}%"),
        ))
    if year and year.isdigit():
        qry = qry.filter(db.extract('year', FacultyPublication.publication_date) == int(year))
    if indexing:
        qry = qry.filter(FacultyPublication.indexing == indexing)

    publications = qry.order_by(FacultyPublication.publication_date.desc().nullslast()).all()
    years    = sorted({p.publication_date.year for p in FacultyPublication.query.all() if p.publication_date}, reverse=True)
    indexings = sorted({p.indexing for p in FacultyPublication.query.all() if p.indexing})

    log_action("Admin viewed publications")

    return render_template(
        "admin/view_publications.html",
        publications=publications,
        years=years,
        indexings=indexings,
        q=q, year=year, indexing=indexing,
    )


# =========================================================
# DOWNLOAD PUBLICATION FILE
# =========================================================

@app.route("/download_publication/<int:pub_id>")
def download_publication(pub_id):

    if session.get("role") != "admin":
        flash("Access denied.", "error")
        return redirect(url_for("login"))

    pub = FacultyPublication.query.get_or_404(pub_id)

    if not pub.article_link:
        flash("No article link attached to this publication.", "error")
        return redirect(url_for("view_publications"))

    log_action(f"Admin accessed publication link id={pub_id}")
    return redirect(pub.article_link)


# =========================================================
# VIEW PROJECTS
# templates/admin/view_projects.html
# =========================================================

@app.route("/view_projects")
def view_projects():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    q            = request.args.get("q", "").strip()
    project_type = request.args.get("project_type", "").strip()
    status       = request.args.get("status", "").strip()

    qry = FacultyProject.query
    if q:
        qry = qry.filter(db.or_(
            FacultyProject.scheme_name.ilike(f"%{q}%"),
            FacultyProject.funding_agency.ilike(f"%{q}%"),
            FacultyProject.pi_co_pi.ilike(f"%{q}%"),
        ))
    if project_type:
        qry = qry.filter(FacultyProject.project_type == project_type)
    if status:
        qry = qry.filter(FacultyProject.status == status)

    projects      = qry.order_by(FacultyProject.created_at.desc()).all()
    project_types = sorted({p.project_type for p in FacultyProject.query.all() if p.project_type})
    statuses      = sorted({p.status for p in FacultyProject.query.all() if p.status})

    log_action("Admin viewed projects")

    return render_template(
        "admin/view_projects.html",
        projects=projects,
        project_types=project_types,
        statuses=statuses,
        q=q, project_type=project_type, status=status,
    )


# =========================================================
# ANALYTICS
# templates/admin/analytics.html
# =========================================================

@app.route("/analytics")
def analytics():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    faculty_count = User.query.filter_by(role="faculty").count()
    profile_count = FacultyProfile.query.count()
    publication_count = FacultyPublication.query.count()
    project_count = FacultyProject.query.count()
    document_count = FacultyDocument.query.count()

    dept_data = db.session.query(
        FacultyProfile.department,
        func.count(FacultyProfile.id)
    ).group_by(FacultyProfile.department).all()

    dept_labels = [d[0] if d[0] else "Unknown" for d in dept_data]
    dept_counts = [d[1] for d in dept_data]

    log_action("Admin viewed analytics")

    return render_template(
        "admin/analytics.html",
        faculty_count=faculty_count,
        profile_count=profile_count,
        publication_count=publication_count,
        project_count=project_count,
        document_count=document_count,
        dept_labels=dept_labels,
        dept_counts=dept_counts
    )


# =========================================================
# VIEW AUDIT LOGS
# templates/admin/view_audit_logs.html
# =========================================================

@app.route("/view_audit_logs")
def view_audit_logs():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    q           = request.args.get("q", "").strip()
    role_filter = request.args.get("role_filter", "").strip()
    date        = request.args.get("date", "").strip()

    qry = AuditLog.query
    if q:
        qry = qry.filter(db.or_(
            AuditLog.action.ilike(f"%{q}%"),
            AuditLog.user_email.ilike(f"%{q}%"),
            AuditLog.ip_address.ilike(f"%{q}%"),
        ))
    if role_filter:
        qry = qry.filter(AuditLog.role == role_filter)
    if date:
        try:
            from datetime import datetime as _dt, timedelta as _td
            d = _dt.strptime(date, "%Y-%m-%d")
            qry = qry.filter(AuditLog.timestamp >= d,
                             AuditLog.timestamp < d + _td(days=1))
        except ValueError:
            pass

    logs = qry.order_by(AuditLog.timestamp.desc()).all()

    log_action("Admin viewed audit logs")

    return render_template(
        "admin/view_audit_logs.html",
        logs=logs,
        q=q, role_filter=role_filter, date=date,
    )


# =========================================================
# REPORT GENERATION
# =========================================================

@app.route("/generate_report")
def generate_report():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    report_path = os.path.join(EXPORT_FOLDER, "faculty_report_all.pdf")
    generate_admin_report(report_path)
    log_action("Admin generated all-faculty report")
    return send_file(report_path, as_attachment=True, download_name="Faculty_Report.pdf")


# =========================================================
# MY REPORT (Faculty downloads their personal PDF report)
# =========================================================

@app.route("/my_report")
def my_report():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    import tempfile, os as _os
    report_path = _os.path.join(EXPORT_FOLDER, f"report_user_{session['user_id']}.pdf")
    generate_personal_report(report_path, session["user_id"])
    log_action("Faculty downloaded personal report")
    return send_file(report_path, as_attachment=True, download_name="My_Academic_Report.pdf")


# =========================================================
# VIEW MY REPORT (Faculty views their personal PDF in browser)
# =========================================================

@app.route("/view_my_report")
def view_my_report():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    import os as _os
    report_path = _os.path.join(EXPORT_FOLDER, f"report_user_{session['user_id']}.pdf")
    generate_personal_report(report_path, session["user_id"])
    log_action("Faculty viewed personal report")
    return send_file(report_path, as_attachment=False, download_name="My_Academic_Report.pdf", mimetype="application/pdf")


# =========================================================
# ADD FACULTY (Admin creates a new faculty user)
# templates/admin/add_faculty.html
# =========================================================

@app.route("/add_faculty", methods=["GET", "POST"])
def add_faculty():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    if request.method == "POST":

        employee_id  = request.form.get("employee_id").strip()
        full_name    = request.form.get("full_name", "").strip()
        email        = request.form.get("email").strip().lower()
        phone        = request.form.get("phone").strip()
        department   = request.form.get("department", "").strip()
        designation  = request.form.get("designation", "").strip()
        password     = request.form.get("password").strip()

        # Password strength check
        if len(password) < 8:
            flash("Password must be at least 8 characters.", "error")
            return redirect(url_for("add_faculty"))

        if not any(c.isupper() for c in password):
            flash("Password must contain at least one uppercase letter.", "error")
            return redirect(url_for("add_faculty"))

        if not any(c.isdigit() for c in password):
            flash("Password must contain at least one number.", "error")
            return redirect(url_for("add_faculty"))

        existing = User.query.filter(
            (User.email == email) | (User.employee_id == employee_id)
        ).first()

        if existing:
            # If the existing user is a faculty with no profile (orphaned from a prior
            # profile-only delete), cascade-delete all their data then allow re-creation
            if existing.role == "faculty" and not FacultyProfile.query.filter_by(user_id=existing.id).first():
                try:
                    AuditLog.query.filter_by(user_id=existing.id).update({"user_id": None})
                    FacultyPublication.query.filter_by(user_id=existing.id).delete()
                    FacultyProject.query.filter_by(user_id=existing.id).delete()
                    FacultyDocument.query.filter_by(user_id=existing.id).delete()
                    FacultyDegree.query.filter_by(user_id=existing.id).delete()
                    FacultyResignation.query.filter_by(user_id=existing.id).delete()
                    db.session.delete(existing)
                    db.session.flush()
                except Exception:
                    db.session.rollback()
                    flash("Could not clean up old account. Please contact admin.", "error")
                    return redirect(url_for("add_faculty"))
            else:
                flash("A user with that email or employee ID already exists.", "error")
                return redirect(url_for("add_faculty"))

        user = User(
            employee_id=employee_id,
            email=email,
            phone=phone,
            password_hash=hash_password(password),
            role="faculty",
            must_reset_password=True
        )

        db.session.add(user)
        db.session.flush()  # get user.id before commit

        # Auto-create a faculty profile so they appear in the profiles list immediately
        profile = FacultyProfile(
            user_id=user.id,
            employee_id=employee_id,
            full_name=full_name or None,
            email_university=email,
            mobile=phone,
            department=department or None,
            designation=designation or None,
        )
        db.session.add(profile)
        db.session.commit()

        log_action(f"Admin added faculty: {email}")

        flash(f"Faculty '{email}' added successfully.", "success")

        return redirect(url_for("view_faculty_profiles"))

    return render_template("admin/add_faculty.html")


# =========================================================
# EDIT FACULTY PROFILE (Admin edits a faculty's profile)
# templates/admin/edit_faculty_profile.html
# =========================================================

@app.route("/edit_faculty_profile/<int:profile_id>", methods=["GET", "POST"])
def edit_faculty_profile(profile_id):

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    profile = FacultyProfile.query.get_or_404(profile_id)

    if request.method == "POST":

        profile.employee_id = request.form.get("employee_id")
        profile.full_name = request.form.get("full_name")
        profile.pan = request.form.get("pan")
        profile.designation = request.form.get("designation")
        profile.department = request.form.get("department")
        profile.qualification = request.form.get("qualification")
        profile.experience_years = request.form.get("experience_years") or None
        profile.mobile = request.form.get("mobile")
        profile.email_university = request.form.get("email_university")
        profile.specialization = request.form.get("specialization")

        db.session.commit()

        log_action(f"Admin edited profile id={profile_id}")

        flash("Profile updated successfully.", "success")

        return redirect(url_for("view_faculty_profiles"))

    return render_template(
        "admin/edit_faculty_profile.html",
        profile=profile
    )


# =========================================================
# DELETE FACULTY PROFILE
# =========================================================

@app.route("/delete_faculty_profile/<int:profile_id>", methods=["POST"])
def delete_faculty_profile(profile_id):

    if session.get("role") != "admin":
        flash("Access denied.", "error")
        return redirect(url_for("login"))

    profile = FacultyProfile.query.get_or_404(profile_id)
    name = profile.full_name or f"id={profile_id}"

    user = User.query.get(profile.user_id)

    if user:
        # Nullify audit log references so history is preserved
        AuditLog.query.filter_by(user_id=user.id).update({"user_id": None})
        # Delete all records linked to this user
        FacultyPublication.query.filter_by(user_id=user.id).delete()
        FacultyProject.query.filter_by(user_id=user.id).delete()
        FacultyDocument.query.filter_by(user_id=user.id).delete()
        FacultyDegree.query.filter_by(user_id=user.id).delete()
        FacultyResignation.query.filter_by(user_id=user.id).delete()

    db.session.delete(profile)
    if user:
        db.session.delete(user)
    db.session.commit()

    log_action(f"Admin deleted faculty user & profile: {name}")
    flash(f"Faculty '{name}' and their account have been deleted.", "success")

    return redirect(url_for("view_faculty_profiles"))


# =========================================================
# ADMIN BULK TEMPLATE DOWNLOADS
# =========================================================

@app.route("/admin_bulk_template/<string:kind>")
def admin_bulk_template(kind):
    if session.get("role") != "admin":
        return redirect(url_for("login"))
    if kind == "faculty_users":
        data, fname = build_faculty_users_template(), "faculty_accounts_template.xlsx"
    elif kind == "faculty_profiles":
        data, fname = build_faculty_profiles_template(), "faculty_profiles_template.xlsx"
    else:
        return "Unknown template type", 404
    return send_file(
        io.BytesIO(data), as_attachment=True, download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================================================
# UPLOAD FACULTY CSV / EXCEL (Bulk create faculty users)
# templates/admin/upload_faculty_csv.html
# =========================================================

@app.route("/upload_faculty_csv", methods=["GET", "POST"])
def upload_faculty_csv():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("upload_faculty_csv"))

        fname = file.filename.lower()
        if not (fname.endswith(".csv") or fname.endswith(".xlsx") or fname.endswith(".xls")):
            flash("Accepted formats: .xlsx, .xls, .csv", "error")
            return redirect(url_for("upload_faculty_csv"))

        file_bytes = file.read()
        records, errors = parse_faculty_users(file_bytes, file.filename)

        added = skipped = 0
        for rec in records:
            existing = User.query.filter(
                (User.email == rec["email"]) | (User.employee_id == rec["employee_id"])
            ).first()
            if existing:
                skipped += 1
                continue
            user = User(
                employee_id   = rec["employee_id"],
                email         = rec["email"],
                phone         = rec["phone"],
                password_hash = hash_password(rec["password"]),
                role          = "faculty",
                must_reset_password = True
            )
            db.session.add(user)
            db.session.flush()
            profile = FacultyProfile(
                user_id          = user.id,
                employee_id      = rec["employee_id"],
                full_name        = rec["full_name"] or None,
                email_university = rec["email"],
                mobile           = rec["phone"],
                department       = rec["department"],
                designation      = rec["designation"],
            )
            db.session.add(profile)
            added += 1

        db.session.commit()
        log_action(f"Admin bulk-uploaded {added} faculty users")

        msg = f"{added} faculty account(s) created."
        if skipped:
            msg += f" {skipped} skipped (already exist)."
        if errors:
            msg += f" {len(errors)} row(s) had errors and were skipped."
        flash(msg, "success" if added else "error")
        return redirect(url_for("view_faculty_profiles"))

    return render_template("admin/upload_faculty_csv.html")


# =========================================================
# UPLOAD FACULTY PROFILES CSV / EXCEL (Bulk upsert profiles)
# templates/admin/upload_faculty_profiles_csv.html
# =========================================================

@app.route("/upload_faculty_profiles_csv", methods=["GET", "POST"])
def upload_faculty_profiles_csv():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("upload_faculty_profiles_csv"))

        fname = file.filename.lower()
        if not (fname.endswith(".csv") or fname.endswith(".xlsx") or fname.endswith(".xls")):
            flash("Accepted formats: .xlsx, .xls, .csv", "error")
            return redirect(url_for("upload_faculty_profiles_csv"))

        file_bytes = file.read()
        records, errors = parse_faculty_profiles(file_bytes, file.filename)

        updated = skipped = 0
        for rec in records:
            user = User.query.filter_by(employee_id=rec["employee_id"]).first()
            if not user:
                skipped += 1
                continue
            profile = FacultyProfile.query.filter_by(user_id=user.id).first()
            if not profile:
                profile = FacultyProfile(user_id=user.id)
            for field in ["employee_id","full_name","pan","designation",
                          "date_of_joining","date_of_birth","appointment_nature",
                          "qualification","department","experience_years",
                          "mobile","email_personal","email_university","specialization"]:
                val = rec.get(field)
                if val is not None:
                    setattr(profile, field, val)
            db.session.add(profile)
            updated += 1

        db.session.commit()
        log_action(f"Admin bulk-updated {updated} faculty profile(s)")

        msg = f"{updated} profile(s) created/updated."
        if skipped:
            msg += f" {skipped} skipped (employee ID not found — create the account first)."
        if errors:
            msg += f" {len(errors)} row(s) had errors."
        flash(msg, "success" if updated else "error")
        return redirect(url_for("view_faculty_profiles"))

    return render_template("admin/upload_faculty_profiles_csv.html")


# =========================================================
# EXPORT FACULTY CSV
# =========================================================

@app.route("/export_faculty_csv")
@app.route("/download_faculty_csv")
def export_faculty_csv_route():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    filepath = os.path.join(EXPORT_FOLDER, "faculty_export.csv")

    export_faculty_csv(filepath)

    log_action("Admin exported faculty CSV")

    return send_file(filepath, as_attachment=True)


# =========================================================
# EXPORT FACULTY PROFILES CSV
# =========================================================

@app.route("/export_faculty_profiles_csv")
@app.route("/download_faculty_profiles_csv")
def export_faculty_profiles_csv_route():

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    filepath = os.path.join(EXPORT_FOLDER, "faculty_profiles.csv")

    export_faculty_profiles_csv(filepath)

    log_action("Admin exported faculty profiles CSV")

    return send_file(filepath, as_attachment=True)


# =========================================================
# FACULTY PROFILE (Faculty views/edits own profile)
# templates/faculty/faculty_profile.html
# =========================================================

@app.route("/faculty_profile", methods=["GET", "POST"])
def faculty_profile():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    user_id = session.get("user_id")

    profile = FacultyProfile.query.filter_by(user_id=user_id).first()
    resignation = FacultyResignation.query.filter_by(user_id=user_id).first()
    degrees = FacultyDegree.query.filter_by(user_id=user_id).all()

    if request.method == "POST":

        # ---- Profile ----
        if not profile:
            profile = FacultyProfile(user_id=user_id)
            db.session.add(profile)

        # employee_id is admin-controlled — faculty cannot change it
        # Only set it on first creation if not already assigned
        if not profile.employee_id:
            user = User.query.get(user_id)
            profile.employee_id = user.employee_id if user else None

        profile.full_name = request.form.get("full_name")
        profile.pan = request.form.get("pan")
        profile.designation = request.form.get("designation")
        profile.qualification = request.form.get("qualification")
        profile.department = request.form.get("department")
        profile.experience_years = request.form.get("experience_years") or None
        profile.appointment_letter_url = request.form.get("appointment_letter_url")
        profile.mobile = request.form.get("mobile")
        profile.email_personal = request.form.get("email_personal")
        profile.email_university = request.form.get("email_university")
        profile.specialization = request.form.get("specialization")
        profile.appointment_nature = request.form.get("appointment_nature")

        doj = request.form.get("date_of_joining")
        dob = request.form.get("date_of_birth")

        if doj:
            try:
                profile.date_of_joining = datetime.strptime(doj, "%Y-%m-%d").date()
            except ValueError:
                pass

        if dob:
            try:
                profile.date_of_birth = datetime.strptime(dob, "%Y-%m-%d").date()
            except ValueError:
                pass

        # ---- Resignation ----
        res_date = request.form.get("resignation_date")
        rel_date = request.form.get("relieving_date")
        reason = request.form.get("reason")
        res_letter = request.form.get("resignation_letter_url")

        if res_date or rel_date or reason:
            if not resignation:
                resignation = FacultyResignation(user_id=user_id)
                db.session.add(resignation)

            if res_date:
                try:
                    resignation.resignation_date = datetime.strptime(res_date, "%Y-%m-%d").date()
                except ValueError:
                    pass

            if rel_date:
                try:
                    resignation.relieving_date = datetime.strptime(rel_date, "%Y-%m-%d").date()
                except ValueError:
                    pass

            resignation.reason = reason
            resignation.resignation_letter_url = res_letter

        # ---- Degree ----
        degree_name = request.form.get("degree_name")

        if degree_name:
            deg_start = request.form.get("degree_start_date")
            deg_award = request.form.get("degree_award_date")
            university = request.form.get("university")

            degree = FacultyDegree(
                user_id=user_id,
                degree_name=degree_name,
                university=university
            )

            if deg_start:
                try:
                    degree.degree_start_date = datetime.strptime(deg_start, "%Y-%m-%d").date()
                except ValueError:
                    pass

            if deg_award:
                try:
                    degree.degree_award_date = datetime.strptime(deg_award, "%Y-%m-%d").date()
                except ValueError:
                    pass

            db.session.add(degree)

        db.session.commit()

        log_action("Faculty updated profile")

        flash("Profile saved successfully.", "success")

        return redirect(url_for("faculty_profile"))

    return render_template(
        "faculty/faculty_profile.html",
        profile=profile,
        resignation=resignation,
        degrees=degrees
    )


# =========================================================
# UPLOAD DOCUMENT (Faculty uploads a document)
# templates/faculty/upload_document.html
# =========================================================

@app.route("/upload_document", methods=["GET", "POST"])
def upload_document():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":

        doc_type = request.form.get("document_type")
        file = request.files.get("file")

        if not file or file.filename == "":
            flash("No file selected.", "error")
            return redirect(url_for("upload_document"))

        if not allowed_file(file.filename):
            flash("File type not allowed. Allowed: PDF, DOC, DOCX, JPG, PNG, XLSX.", "error")
            return redirect(url_for("upload_document"))

        user_folder = os.path.join(UPLOAD_FOLDER, f"user_{session['user_id']}")
        os.makedirs(user_folder, exist_ok=True)
        filename = f"{uuid.uuid4().hex}_{secure_filename(file.filename)}"
        save_path = os.path.join(user_folder, filename)
        file.save(save_path)

        doc = FacultyDocument(
            user_id=session["user_id"],
            document_type=doc_type,
            file_name=filename,
            file_path=save_path
        )

        db.session.add(doc)
        db.session.commit()

        log_action(f"Faculty uploaded document: {filename}")

        flash("Document uploaded successfully.", "success")

        return redirect(url_for("my_documents"))

    return render_template("faculty/upload_document.html")


# =========================================================
# MY DOCUMENTS (Faculty views own documents)
# templates/faculty/my_documents.html
# =========================================================

@app.route("/my_documents")
def my_documents():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    documents = FacultyDocument.query.filter_by(
        user_id=session["user_id"]
    ).order_by(FacultyDocument.uploaded_at.desc()).all()

    log_action("Faculty viewed their documents")

    return render_template(
        "faculty/my_documents.html",
        documents=documents
    )


# =========================================================
# DOWNLOAD DOCUMENT
# =========================================================

@app.route("/download_document/<int:doc_id>")
def download_document(doc_id):

    if not session.get("user_id"):
        return redirect(url_for("login"))

    doc = FacultyDocument.query.get_or_404(doc_id)

    # Faculty can only download their own; admin can download any
    if session.get("role") == "faculty" and doc.user_id != session["user_id"]:
        flash("Access denied.", "error")
        return redirect(url_for("my_documents"))

    # Path traversal guard — ensure file is inside uploads folder
    real_path = os.path.realpath(doc.file_path)
    real_upload = os.path.realpath(UPLOAD_FOLDER)
    if not real_path.startswith(real_upload + os.sep) and real_path != real_upload:
        log_action(f"Blocked suspicious download path for doc id={doc_id}")
        flash("Invalid file path.", "error")
        return redirect(url_for("my_documents"))

    if not os.path.exists(real_path):
        flash("File not found.", "error")
        return redirect(url_for("my_documents"))

    log_action(f"Downloaded document id={doc_id}")

    return send_file(real_path, as_attachment=True)


# =========================================================
# ADD PUBLICATION (Faculty adds a publication)
# templates/faculty/add_publication.html
# =========================================================

@app.route("/api/fetch_doi")
def fetch_doi():
    """Fetch publication metadata from CrossRef using a DOI."""
    import requests as _req
    doi = request.args.get("doi", "").strip().lstrip("https://doi.org/").lstrip("http://dx.doi.org/")
    if not doi:
        return {"error": "No DOI provided"}, 400
    try:
        r = _req.get(
            f"https://api.crossref.org/works/{doi}",
            headers={"User-Agent": "FacultyMIS/1.0 (mailto:admin@university.edu)"},
            timeout=8
        )
        if r.status_code != 200:
            return {"error": "DOI not found"}, 404
        msg = r.json().get("message", {})
        authors = ", ".join(
            f"{a.get('given', '')} {a.get('family', '')}".strip()
            for a in msg.get("author", [])
        )
        year = None
        dp = msg.get("published", msg.get("published-print", msg.get("published-online", {})))
        parts = dp.get("date-parts", [[None]])[0]
        if parts:
            year = parts[0]
        journal = (msg.get("container-title") or [""])[0]
        title   = (msg.get("title") or [""])[0]

        # Extract ISSN
        issn_list = msg.get("ISSN") or []
        issn_isbn = issn_list[0] if issn_list else ""

        # Extract ISBN (for books / book chapters)
        isbn_list = msg.get("ISBN") or []
        isbn = isbn_list[0] if isbn_list else ""

        # Extract publication date as DD-MMM-YYYY
        pub_date_str = ""
        if parts and len(parts) >= 3:
            try:
                from datetime import date as _date
                pub_date_str = _date(parts[0], parts[1], parts[2]).strftime("%d-%b-%Y")
            except Exception:
                pass
        elif parts and len(parts) >= 1 and year:
            pub_date_str = f"01-Jan-{year}"

        return {
            "title":       title,
            "authors":     authors,
            "journal":     journal,
            "year":        year,
            "publisher":   msg.get("publisher", ""),
            "volume":      msg.get("volume", ""),
            "issue":       msg.get("issue", ""),
            "pages":       msg.get("page", ""),
            "doi":         doi,
            "issn_isbn":   issn_isbn,
            "isbn":        isbn,
            "pub_date":    pub_date_str,
            "type":        msg.get("type", ""),
        }
    except Exception as e:
        return {"error": str(e)}, 500


@app.route("/add_publication", methods=["GET", "POST"])
def add_publication():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":

        title            = request.form.get("title")
        author_position  = request.form.get("author_position") or None
        scholar_id       = request.form.get("scholar_id") or None
        journal          = request.form.get("journal") or None
        pub_date_raw     = request.form.get("publication_date") or None
        issn_isbn        = request.form.get("issn_isbn") or None
        h_index          = request.form.get("h_index") or None
        citation_index   = request.form.get("citation_index") or None
        journal_quartile = request.form.get("journal_quartile") or None
        publication_type = request.form.get("publication_type") or None
        impact_factor    = request.form.get("impact_factor") or None
        indexing         = request.form.get("indexing") or None
        doi              = request.form.get("doi") or None
        article_link     = request.form.get("article_link") or None

        # Parse publication date
        publication_date = None
        if pub_date_raw:
            for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d-%m-%Y", "%d/%m/%Y"):
                try:
                    publication_date = datetime.strptime(pub_date_raw.strip(), fmt).date()
                    break
                except ValueError:
                    pass

        pub = FacultyPublication(
            user_id          = session["user_id"],
            title            = title,
            author_position  = author_position,
            scholar_id       = scholar_id,
            journal          = journal,
            publication_date = publication_date,
            issn_isbn        = issn_isbn,
            h_index          = h_index,
            citation_index   = citation_index,
            journal_quartile = journal_quartile,
            publication_type = publication_type,
            impact_factor    = impact_factor,
            indexing         = indexing,
            doi              = doi,
            article_link     = article_link,
        )

        db.session.add(pub)
        db.session.commit()

        log_action(f"Faculty added publication: {title}")

        flash("Publication added successfully.", "success")

        return redirect(url_for("my_publications"))

    return render_template("faculty/add_publication.html")


# =========================================================
# MY PUBLICATIONS (Faculty views own publications)
# templates/faculty/my_publications.html
# =========================================================

@app.route("/my_publications")
def my_publications():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    publications = FacultyPublication.query.filter_by(
        user_id=session["user_id"]
    ).order_by(FacultyPublication.created_at.desc()).all()

    log_action("Faculty viewed their publications")

    return render_template(
        "faculty/my_publications.html",
        publications=publications
    )


# =========================================================
# ADD PROJECT (Faculty adds a research project)
# templates/faculty/add_project.html
# =========================================================

@app.route("/add_project", methods=["GET", "POST"])
def add_project():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":

        scheme_name               = request.form.get("scheme_name", "").strip()
        pi_co_pi                  = request.form.get("pi_co_pi", "").strip() or None
        funding_agency            = request.form.get("funding_agency", "").strip() or None
        project_type              = request.form.get("project_type") or None
        department                = request.form.get("department", "").strip() or None
        date_of_award_raw         = request.form.get("date_of_award")
        amount                    = request.form.get("amount") or None
        duration_years            = request.form.get("duration_years", "").strip() or None
        status                    = request.form.get("status") or None
        date_of_completion_raw    = request.form.get("date_of_completion")
        objectives                = request.form.get("objectives", "").strip() or None
        collaborating_institutions = request.form.get("collaborating_institutions", "").strip() or None
        document_link             = request.form.get("document_link", "").strip() or None

        if not scheme_name:
            flash("Scheme / Project name is required.", "error")
            return redirect(url_for("add_project"))

        project = FacultyProject(
            user_id=session["user_id"],
            scheme_name=scheme_name,
            pi_co_pi=pi_co_pi,
            funding_agency=funding_agency,
            project_type=project_type,
            department=department,
            amount=float(amount) if amount else None,
            duration_years=duration_years,
            status=status,
            objectives=objectives,
            collaborating_institutions=collaborating_institutions,
            document_link=document_link,
        )

        if date_of_award_raw:
            try:
                project.date_of_award = datetime.strptime(date_of_award_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        if date_of_completion_raw:
            try:
                project.date_of_completion = datetime.strptime(date_of_completion_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        db.session.add(project)
        db.session.commit()

        log_action(f"Faculty added project: {scheme_name}")
        flash("Project added successfully.", "success")
        return redirect(url_for("my_projects"))

    return render_template("faculty/add_project.html")


# =========================================================
# MY PROJECTS (Faculty views own projects)
# templates/faculty/my_projects.html
# =========================================================

@app.route("/my_projects")
def my_projects():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    projects = FacultyProject.query.filter_by(
        user_id=session["user_id"]
    ).order_by(FacultyProject.created_at.desc()).all()

    log_action("Faculty viewed their projects")

    return render_template(
        "faculty/my_projects.html",
        projects=projects
    )


# =========================================================
# EDIT PUBLICATION (Faculty edits their own publication)
# templates/faculty/edit_publication.html
# =========================================================

@app.route("/edit_publication/<int:pub_id>", methods=["GET", "POST"])
def edit_publication(pub_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    pub = FacultyPublication.query.filter_by(
        id=pub_id, user_id=session["user_id"]
    ).first_or_404()

    if request.method == "POST":

        pub.title            = request.form.get("title", pub.title)
        pub.author_position  = request.form.get("author_position") or None
        pub.scholar_id       = request.form.get("scholar_id") or None
        pub.journal          = request.form.get("journal") or None
        pub.issn_isbn        = request.form.get("issn_isbn") or None
        pub.h_index          = request.form.get("h_index") or None
        pub.citation_index   = request.form.get("citation_index") or None
        pub.journal_quartile = request.form.get("journal_quartile") or None
        pub.publication_type = request.form.get("publication_type") or None
        pub.impact_factor    = request.form.get("impact_factor") or None
        pub.indexing         = request.form.get("indexing") or None
        pub.doi              = request.form.get("doi") or None
        pub.article_link     = request.form.get("article_link") or None

        pub_date_raw = request.form.get("publication_date") or None
        pub.publication_date = None
        if pub_date_raw:
            for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d-%m-%Y", "%d/%m/%Y"):
                try:
                    pub.publication_date = datetime.strptime(pub_date_raw.strip(), fmt).date()
                    break
                except ValueError:
                    pass

        db.session.commit()
        log_action(f"Faculty edited publication: {pub.title}")
        flash("Publication updated successfully.", "success")
        return redirect(url_for("my_publications"))

    return render_template("faculty/edit_publication.html", pub=pub)


# =========================================================
# EDIT PROJECT (Faculty edits their own project)
# templates/faculty/edit_project.html
# =========================================================

@app.route("/edit_project/<int:project_id>", methods=["GET", "POST"])
def edit_project(project_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    project = FacultyProject.query.filter_by(
        id=project_id, user_id=session["user_id"]
    ).first_or_404()

    if request.method == "POST":

        scheme_name = request.form.get("scheme_name", "").strip()
        if not scheme_name:
            flash("Scheme / Project name is required.", "error")
            return redirect(url_for("edit_project", project_id=project_id))

        project.scheme_name                = scheme_name
        project.pi_co_pi                   = request.form.get("pi_co_pi", "").strip() or None
        project.funding_agency             = request.form.get("funding_agency", "").strip() or None
        project.project_type               = request.form.get("project_type") or None
        project.department                 = request.form.get("department", "").strip() or None
        amount                             = request.form.get("amount") or None
        project.amount                     = float(amount) if amount else None
        project.duration_years             = request.form.get("duration_years", "").strip() or None
        project.status                     = request.form.get("status") or None
        project.objectives                 = request.form.get("objectives", "").strip() or None
        project.collaborating_institutions = request.form.get("collaborating_institutions", "").strip() or None
        project.document_link              = request.form.get("document_link", "").strip() or None

        date_of_award_raw      = request.form.get("date_of_award")
        date_of_completion_raw = request.form.get("date_of_completion")
        if date_of_award_raw:
            try:
                project.date_of_award = datetime.strptime(date_of_award_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        else:
            project.date_of_award = None
        if date_of_completion_raw:
            try:
                project.date_of_completion = datetime.strptime(date_of_completion_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        else:
            project.date_of_completion = None

        db.session.commit()
        log_action(f"Faculty edited project: {project.scheme_name}")
        flash("Project updated successfully.", "success")
        return redirect(url_for("my_projects"))

    return render_template("faculty/edit_project.html", project=project)


# =========================================================
# DELETE PUBLICATION (Faculty deletes their own publication)
# =========================================================

@app.route("/delete_publication/<int:pub_id>", methods=["POST"])
def delete_publication(pub_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    pub = FacultyPublication.query.filter_by(
        id=pub_id, user_id=session["user_id"]
    ).first_or_404()

    db.session.delete(pub)
    db.session.commit()
    log_action(f"Faculty deleted publication: {pub.title}")
    flash("Publication deleted.", "success")
    return redirect(url_for("my_publications"))


# =========================================================
# DELETE PROJECT (Faculty deletes their own project)
# =========================================================

@app.route("/delete_project/<int:project_id>", methods=["POST"])
def delete_project(project_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    project = FacultyProject.query.filter_by(
        id=project_id, user_id=session["user_id"]
    ).first_or_404()

    db.session.delete(project)
    db.session.commit()
    log_action(f"Faculty deleted project: {project.title}")
    flash("Project deleted.", "success")
    return redirect(url_for("my_projects"))


# =========================================================
# VALUE ADDED PROGRAMS — LIST (courses attended + offered)
# templates/faculty/my_courses.html
# =========================================================

@app.route("/my_courses")
def my_courses():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    attended = CourseAttended.query.filter_by(
        user_id=session["user_id"]
    ).order_by(CourseAttended.created_at.desc()).all()

    offered = CourseOffered.query.filter_by(
        user_id=session["user_id"]
    ).order_by(CourseOffered.created_at.desc()).all()

    log_action("Faculty viewed Value Added Programs")

    return render_template(
        "faculty/my_courses.html",
        attended=attended,
        offered=offered
    )


# =========================================================
# ADD COURSE ATTENDED
# templates/faculty/add_course_attended.html
# =========================================================

@app.route("/add_course_attended", methods=["GET", "POST"])
def add_course_attended():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":

        course_name     = request.form.get("course_name")
        online_course   = request.form.get("online_course_name")
        date_from_raw   = request.form.get("date_from")
        date_to_raw     = request.form.get("date_to")
        mode            = request.form.get("mode")
        contact_hours   = request.form.get("contact_hours") or None
        offered_by      = request.form.get("offered_by")
        cert_link       = request.form.get("certificate_link")

        entry = CourseAttended(
            user_id=session["user_id"],
            course_name=course_name,
            online_course_name=online_course,
            mode=mode,
            contact_hours=float(contact_hours) if contact_hours else None,
            offered_by=offered_by,
            certificate_link=cert_link
        )

        if date_from_raw:
            try:
                entry.date_from = datetime.strptime(date_from_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        if date_to_raw:
            try:
                entry.date_to = datetime.strptime(date_to_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        db.session.add(entry)
        db.session.commit()
        log_action(f"Faculty added course attended: {course_name}")
        flash("Course attended added successfully.", "success")
        return redirect(url_for("my_courses"))

    return render_template("faculty/add_course_attended.html")


# =========================================================
# EDIT COURSE ATTENDED
# templates/faculty/edit_course_attended.html
# =========================================================

@app.route("/edit_course_attended/<int:entry_id>", methods=["GET", "POST"])
def edit_course_attended(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = CourseAttended.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    if request.method == "POST":

        entry.course_name        = request.form.get("course_name")
        entry.online_course_name = request.form.get("online_course_name")
        entry.mode               = request.form.get("mode")
        ch = request.form.get("contact_hours")
        entry.contact_hours      = float(ch) if ch else None
        entry.offered_by         = request.form.get("offered_by")
        entry.certificate_link   = request.form.get("certificate_link")

        date_from_raw = request.form.get("date_from")
        date_to_raw   = request.form.get("date_to")
        if date_from_raw:
            try:
                entry.date_from = datetime.strptime(date_from_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        if date_to_raw:
            try:
                entry.date_to = datetime.strptime(date_to_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        db.session.commit()
        log_action(f"Faculty edited course attended: {entry.course_name}")
        flash("Course updated.", "success")
        return redirect(url_for("my_courses"))

    return render_template("faculty/edit_course_attended.html", entry=entry)


# =========================================================
# DELETE COURSE ATTENDED
# =========================================================

@app.route("/delete_course_attended/<int:entry_id>", methods=["POST"])
def delete_course_attended(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = CourseAttended.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    db.session.delete(entry)
    db.session.commit()
    log_action(f"Faculty deleted course attended: {entry.course_name}")
    flash("Entry deleted.", "success")
    return redirect(url_for("my_courses"))


# =========================================================
# ADD COURSE OFFERED
# templates/faculty/add_course_offered.html
# =========================================================

@app.route("/add_course_offered", methods=["GET", "POST"])
def add_course_offered():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":

        course_name         = request.form.get("course_name")
        online_course       = request.form.get("online_course_name")
        credits_assigned    = request.form.get("credits_assigned")
        program_name        = request.form.get("program_name")
        department          = request.form.get("department")
        date_from_raw       = request.form.get("date_from")
        date_to_raw         = request.form.get("date_to")
        times_offered       = request.form.get("times_offered") or None
        mode                = request.form.get("mode")
        contact_hours       = request.form.get("contact_hours") or None
        students_enrolled   = request.form.get("students_enrolled_link")
        students_completing = request.form.get("students_completing_link")
        attendance_link     = request.form.get("attendance_link")
        brochure_link       = request.form.get("brochure_link")
        cert_link           = request.form.get("certificate_link")

        entry = CourseOffered(
            user_id=session["user_id"],
            course_name=course_name,
            online_course_name=online_course,
            credits_assigned=credits_assigned,
            program_name=program_name,
            department=department,
            times_offered=int(times_offered) if times_offered else None,
            mode=mode,
            contact_hours=float(contact_hours) if contact_hours else None,
            students_enrolled_link=students_enrolled,
            students_completing_link=students_completing,
            attendance_link=attendance_link,
            brochure_link=brochure_link,
            certificate_link=cert_link
        )

        if date_from_raw:
            try:
                entry.date_from = datetime.strptime(date_from_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        if date_to_raw:
            try:
                entry.date_to = datetime.strptime(date_to_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        db.session.add(entry)
        db.session.commit()
        log_action(f"Faculty added course offered: {course_name}")
        flash("Course offered added successfully.", "success")
        return redirect(url_for("my_courses"))

    return render_template("faculty/add_course_offered.html")


# =========================================================
# EDIT COURSE OFFERED
# templates/faculty/edit_course_offered.html
# =========================================================

@app.route("/edit_course_offered/<int:entry_id>", methods=["GET", "POST"])
def edit_course_offered(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = CourseOffered.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    if request.method == "POST":

        entry.course_name            = request.form.get("course_name")
        entry.online_course_name     = request.form.get("online_course_name")
        entry.credits_assigned       = request.form.get("credits_assigned")
        entry.program_name           = request.form.get("program_name")
        entry.department             = request.form.get("department")
        to = request.form.get("times_offered")
        entry.times_offered          = int(to) if to else None
        entry.mode                   = request.form.get("mode")
        ch = request.form.get("contact_hours")
        entry.contact_hours          = float(ch) if ch else None
        entry.students_enrolled_link  = request.form.get("students_enrolled_link")
        entry.students_completing_link = request.form.get("students_completing_link")
        entry.attendance_link        = request.form.get("attendance_link")
        entry.brochure_link          = request.form.get("brochure_link")
        entry.certificate_link       = request.form.get("certificate_link")

        date_from_raw = request.form.get("date_from")
        date_to_raw   = request.form.get("date_to")
        if date_from_raw:
            try:
                entry.date_from = datetime.strptime(date_from_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        if date_to_raw:
            try:
                entry.date_to = datetime.strptime(date_to_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        db.session.commit()
        log_action(f"Faculty edited course offered: {entry.course_name}")
        flash("Course updated.", "success")
        return redirect(url_for("my_courses"))

    return render_template("faculty/edit_course_offered.html", entry=entry)


# =========================================================
# DELETE COURSE OFFERED
# =========================================================

@app.route("/delete_course_offered/<int:entry_id>", methods=["POST"])
def delete_course_offered(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = CourseOffered.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    db.session.delete(entry)
    db.session.commit()
    log_action(f"Faculty deleted course offered: {entry.course_name}")
    flash("Entry deleted.", "success")
    return redirect(url_for("my_courses"))


# =========================================================
# BULK UPLOAD — TEMPLATE DOWNLOADS
# =========================================================

@app.route("/bulk_template/<string:kind>")
def bulk_template(kind):
    """Download a pre-built Excel template (publications or projects)."""
    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if kind == "publications":
        data     = build_publications_template()
        filename = "publications_template.xlsx"
    elif kind == "projects":
        data     = build_projects_template()
        filename = "projects_template.xlsx"
    elif kind == "courses_attended":
        data     = build_courses_attended_template()
        filename = "courses_attended_template.xlsx"
    elif kind == "courses_offered":
        data     = build_courses_offered_template()
        filename = "courses_offered_template.xlsx"
    elif kind == "awards":
        data     = build_awards_template()
        filename = "awards_template.xlsx"
    elif kind == "book_chapters":
        data     = build_book_chapters_template()
        filename = "book_chapters_template.xlsx"
    elif kind == "guest_lectures":
        data     = build_guest_lectures_template()
        filename = "guest_lectures_template.xlsx"
    elif kind == "patents":
        data     = build_patents_template()
        filename = "patents_template.xlsx"
    elif kind == "fellowships":
        data     = build_fellowships_template()
        filename = "fellowships_template.xlsx"
    elif kind == "conferences_participated":
        data     = build_conferences_participated_template()
        filename = "conferences_participated_template.xlsx"
    elif kind == "conferences_organised":
        data     = build_conferences_organised_template()
        filename = "conferences_organised_template.xlsx"
    elif kind == "fdp_participated":
        data     = build_fdp_participated_template()
        filename = "fdp_participated_template.xlsx"
    elif kind == "fdp_organised":
        data     = build_fdp_organised_template()
        filename = "fdp_organised_template.xlsx"
    else:
        return "Unknown template type", 404

    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================================================
# BULK UPLOAD — PUBLICATIONS
# templates/faculty/bulk_upload_publications.html
# =========================================================

@app.route("/bulk_upload_publications", methods=["GET", "POST"])
def bulk_upload_publications():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("bulk_upload_publications"))

        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("csv", "xlsx", "xls"):
            flash("Only CSV or Excel (.xlsx/.xls) files are supported.", "error")
            return redirect(url_for("bulk_upload_publications"))

        file_bytes = file.read()
        records, errors = parse_publications(file_bytes, file.filename)

        inserted = 0
        for rec in records:
            pub = FacultyPublication(
                user_id          = session["user_id"],
                title            = rec["title"],
                author_position  = rec.get("author_position"),
                scholar_id       = rec.get("scholar_id"),
                journal          = rec.get("journal"),
                publication_date = rec.get("publication_date"),
                issn_isbn        = rec.get("issn_isbn"),
                h_index          = rec.get("h_index"),
                citation_index   = rec.get("citation_index"),
                journal_quartile = rec.get("journal_quartile"),
                publication_type = rec.get("publication_type"),
                impact_factor    = rec.get("impact_factor"),
                indexing         = rec.get("indexing"),
                doi              = rec.get("doi"),
                article_link     = rec.get("article_link"),
            )
            db.session.add(pub)
            inserted += 1

        db.session.commit()
        log_action(f"Faculty bulk-uploaded {inserted} publication(s)")

        if errors:
            skipped = ", ".join(f"row {r}: {m}" for r, m in errors[:5])
            flash(f"{inserted} publication(s) imported. Skipped rows — {skipped}", "warning")
        else:
            flash(f"{inserted} publication(s) imported successfully.", "success")

        return redirect(url_for("my_publications"))

    return render_template("faculty/bulk_upload_publications.html")


# =========================================================
# BULK UPLOAD — PROJECTS
# templates/faculty/bulk_upload_projects.html
# =========================================================

@app.route("/bulk_upload_projects", methods=["GET", "POST"])
def bulk_upload_projects():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("bulk_upload_projects"))

        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("csv", "xlsx", "xls"):
            flash("Only CSV or Excel (.xlsx/.xls) files are supported.", "error")
            return redirect(url_for("bulk_upload_projects"))

        file_bytes = file.read()
        records, errors = parse_projects(file_bytes, file.filename)

        inserted = 0
        for rec in records:
            proj = FacultyProject(
                user_id                      = session["user_id"],
                scheme_name                  = rec["scheme_name"],
                pi_co_pi                     = rec["pi_co_pi"],
                funding_agency               = rec["funding_agency"],
                project_type                 = rec["project_type"],
                department                   = rec["department"],
                date_of_award                = rec["date_of_award"],
                amount                       = rec["amount"],
                duration_years               = rec["duration_years"],
                status                       = rec["status"],
                date_of_completion           = rec["date_of_completion"],
                objectives                   = rec["objectives"],
                collaborating_institutions   = rec["collaborating_institutions"],
                document_link                = rec["document_link"],
            )
            db.session.add(proj)
            inserted += 1

        db.session.commit()
        log_action(f"Faculty bulk-uploaded {inserted} project(s)")

        if errors:
            skipped = ", ".join(f"row {r}: {m}" for r, m in errors[:5])
            flash(f"{inserted} project(s) imported. Skipped rows — {skipped}", "warning")
        else:
            flash(f"{inserted} project(s) imported successfully.", "success")

        return redirect(url_for("my_projects"))

    return render_template("faculty/bulk_upload_projects.html")


# =========================================================
# BULK UPLOAD — COURSES ATTENDED
# templates/faculty/bulk_upload_courses_attended.html
# =========================================================

@app.route("/bulk_upload_courses_attended", methods=["GET", "POST"])
def bulk_upload_courses_attended():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("bulk_upload_courses_attended"))

        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("csv", "xlsx", "xls"):
            flash("Only CSV or Excel (.xlsx/.xls) files are supported.", "error")
            return redirect(url_for("bulk_upload_courses_attended"))

        file_bytes = file.read()
        records, errors = parse_courses_attended(file_bytes, file.filename)

        inserted = 0
        for rec in records:
            entry = CourseAttended(
                user_id            = session["user_id"],
                course_name        = rec["course_name"],
                online_course_name = rec["online_course_name"],
                date_from          = rec["date_from"],
                date_to            = rec["date_to"],
                mode               = rec["mode"],
                contact_hours      = rec["contact_hours"],
                offered_by         = rec["offered_by"],
                certificate_link   = rec["certificate_link"],
            )
            db.session.add(entry)
            inserted += 1

        db.session.commit()
        log_action(f"Faculty bulk-uploaded {inserted} course(s) attended")

        if errors:
            skipped = ", ".join(f"row {r}: {m}" for r, m in errors[:5])
            flash(f"{inserted} record(s) imported. Skipped rows — {skipped}", "warning")
        else:
            flash(f"{inserted} course(s) attended imported successfully.", "success")

        return redirect(url_for("my_courses"))

    return render_template("faculty/bulk_upload_courses_attended.html")


# =========================================================
# BULK UPLOAD — COURSES OFFERED
# templates/faculty/bulk_upload_courses_offered.html
# =========================================================

@app.route("/bulk_upload_courses_offered", methods=["GET", "POST"])
def bulk_upload_courses_offered():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("bulk_upload_courses_offered"))

        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("csv", "xlsx", "xls"):
            flash("Only CSV or Excel (.xlsx/.xls) files are supported.", "error")
            return redirect(url_for("bulk_upload_courses_offered"))

        file_bytes = file.read()
        records, errors = parse_courses_offered(file_bytes, file.filename)

        inserted = 0
        for rec in records:
            entry = CourseOffered(
                user_id                  = session["user_id"],
                course_name              = rec["course_name"],
                online_course_name       = rec["online_course_name"],
                credits_assigned         = rec["credits_assigned"],
                program_name             = rec["program_name"],
                department               = rec["department"],
                date_from                = rec["date_from"],
                date_to                  = rec["date_to"],
                times_offered            = rec["times_offered"],
                mode                     = rec["mode"],
                contact_hours            = rec["contact_hours"],
                students_enrolled_link   = rec["students_enrolled_link"],
                students_completing_link = rec["students_completing_link"],
                attendance_link          = rec["attendance_link"],
                brochure_link            = rec["brochure_link"],
                certificate_link         = rec["certificate_link"],
            )
            db.session.add(entry)
            inserted += 1

        db.session.commit()
        log_action(f"Faculty bulk-uploaded {inserted} course(s) offered")

        if errors:
            skipped = ", ".join(f"row {r}: {m}" for r, m in errors[:5])
            flash(f"{inserted} record(s) imported. Skipped rows — {skipped}", "warning")
        else:
            flash(f"{inserted} course(s) offered imported successfully.", "success")

        return redirect(url_for("my_courses"))

    return render_template("faculty/bulk_upload_courses_offered.html")


# =========================================================
# MY AWARDS
# templates/faculty/my_awards.html
# =========================================================

@app.route("/my_awards")
def my_awards():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()
    query = FacultyAward.query.filter_by(user_id=session["user_id"])
    if q:
        query = query.filter(
            FacultyAward.title.ilike(f"%{q}%") |
            FacultyAward.awarding_agency.ilike(f"%{q}%") |
            FacultyAward.nature_of_award.ilike(f"%{q}%")
        )
    awards = query.order_by(FacultyAward.date_of_award.desc()).all()
    return render_template("faculty/my_awards.html", awards=awards, q=q)


# =========================================================
# ADD AWARD
# templates/faculty/add_award.html
# =========================================================

@app.route("/add_award", methods=["GET", "POST"])
def add_award():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    if request.method == "POST":
        title = request.form.get("title", "").strip()
        if not title:
            flash("Title is required.", "error")
            return redirect(url_for("add_award"))
        award = FacultyAward(
            user_id         = session["user_id"],
            title           = title,
            nature_of_award = request.form.get("nature_of_award", "").strip() or None,
            event_level     = request.form.get("event_level", "").strip() or None,
            date_of_award   = (datetime.strptime(request.form["date_of_award"], "%Y-%m-%d").date()
                               if request.form.get("date_of_award") else None),
            category        = request.form.get("category", "").strip() or None,
            awarding_agency = request.form.get("awarding_agency", "").strip() or None,
            award_amount    = request.form.get("award_amount", "").strip() or None,
            research_area   = request.form.get("research_area", "").strip() or None,
            collaborators   = request.form.get("collaborators", "").strip() or None,
            document_link   = request.form.get("document_link", "").strip() or None,
        )
        db.session.add(award)
        db.session.commit()
        log_action(f"Faculty added award: {title}")
        flash("Award added successfully.", "success")
        return redirect(url_for("my_awards"))
    return render_template("faculty/add_award.html")


# =========================================================
# EDIT AWARD
# templates/faculty/edit_award.html
# =========================================================

@app.route("/edit_award/<int:award_id>", methods=["GET", "POST"])
def edit_award(award_id):
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    award = FacultyAward.query.filter_by(id=award_id, user_id=session["user_id"]).first_or_404()
    if request.method == "POST":
        title = request.form.get("title", "").strip()
        if not title:
            flash("Title is required.", "error")
            return redirect(url_for("edit_award", award_id=award_id))
        award.title           = title
        award.nature_of_award = request.form.get("nature_of_award", "").strip() or None
        award.event_level     = request.form.get("event_level", "").strip() or None
        award.date_of_award   = (datetime.strptime(request.form["date_of_award"], "%Y-%m-%d").date()
                                 if request.form.get("date_of_award") else None)
        award.category        = request.form.get("category", "").strip() or None
        award.awarding_agency = request.form.get("awarding_agency", "").strip() or None
        award.award_amount    = request.form.get("award_amount", "").strip() or None
        award.research_area   = request.form.get("research_area", "").strip() or None
        award.collaborators   = request.form.get("collaborators", "").strip() or None
        award.document_link   = request.form.get("document_link", "").strip() or None
        db.session.commit()
        log_action(f"Faculty edited award #{award_id}")
        flash("Award updated successfully.", "success")
        return redirect(url_for("my_awards"))
    return render_template("faculty/edit_award.html", award=award)


# =========================================================
# DELETE AWARD
# =========================================================

@app.route("/delete_award/<int:award_id>", methods=["POST"])
def delete_award(award_id):
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    award = FacultyAward.query.filter_by(id=award_id, user_id=session["user_id"]).first_or_404()
    db.session.delete(award)
    db.session.commit()
    log_action(f"Faculty deleted award #{award_id}")
    flash("Award deleted.", "success")
    return redirect(url_for("my_awards"))


# =========================================================
# BULK UPLOAD AWARDS
# templates/faculty/bulk_upload_awards.html
# =========================================================

@app.route("/bulk_upload_awards", methods=["GET", "POST"])
def bulk_upload_awards():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    if request.method == "POST":
        if request.form.get("download_template"):
            buf = build_awards_template()
            return send_file(
                io.BytesIO(buf),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="awards_template.xlsx"
            )
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("bulk_upload_awards"))
        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("csv", "xlsx", "xls"):
            flash("Only CSV or Excel (.xlsx/.xls) files are supported.", "error")
            return redirect(url_for("bulk_upload_awards"))
        file_bytes = file.read()
        records, errors = parse_awards(file_bytes, file.filename)
        inserted = 0
        for rec in records:
            entry = FacultyAward(
                user_id         = session["user_id"],
                title           = rec["title"],
                nature_of_award = rec["nature_of_award"],
                event_level     = rec["event_level"],
                date_of_award   = rec["date_of_award"],
                category        = rec["category"],
                awarding_agency = rec["awarding_agency"],
                award_amount    = rec["award_amount"],
                research_area   = rec["research_area"],
                collaborators   = rec["collaborators"],
                document_link   = rec["document_link"],
            )
            db.session.add(entry)
            inserted += 1
        db.session.commit()
        log_action(f"Faculty bulk-uploaded {inserted} award(s)")
        if errors:
            skipped = ", ".join(f"row {r}: {m}" for r, m in errors[:5])
            flash(f"{inserted} record(s) imported. Skipped rows — {skipped}", "warning")
        else:
            flash(f"{inserted} award(s) imported successfully.", "success")
        return redirect(url_for("my_awards"))
    return render_template("faculty/bulk_upload_awards.html")


# =========================================================
# ADMIN — VIEW ALL AWARDS
# templates/admin/view_awards.html
# =========================================================

@app.route("/admin/view_awards")
def view_awards():
    if session.get("role") != "admin":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()
    query = FacultyAward.query
    if q:
        query = query.filter(
            FacultyAward.title.ilike(f"%{q}%") |
            FacultyAward.awarding_agency.ilike(f"%{q}%") |
            FacultyAward.nature_of_award.ilike(f"%{q}%")
        )
    awards = query.order_by(FacultyAward.date_of_award.desc()).all()
    return render_template("admin/view_awards.html", awards=awards, q=q)


# =========================================================
# MY BOOK CHAPTERS
# templates/faculty/my_book_chapters.html
# =========================================================

@app.route("/my_book_chapters")
def my_book_chapters():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()
    query = FacultyBookChapter.query.filter_by(user_id=session["user_id"])
    if q:
        query = query.filter(
            FacultyBookChapter.book_title.ilike(f"%{q}%") |
            FacultyBookChapter.chapter_title.ilike(f"%{q}%") |
            FacultyBookChapter.co_authors.ilike(f"%{q}%")
        )
    chapters = query.order_by(FacultyBookChapter.publication_date.desc()).all()
    return render_template("faculty/my_book_chapters.html", chapters=chapters, q=q)


# =========================================================
# ADD BOOK CHAPTER
# templates/faculty/add_book_chapter.html
# =========================================================

@app.route("/add_book_chapter", methods=["GET", "POST"])
def add_book_chapter():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    if request.method == "POST":
        book_title = request.form.get("book_title", "").strip()
        if not book_title:
            flash("Title of the Book Published is required.", "error")
            return redirect(url_for("add_book_chapter"))
        entry = FacultyBookChapter(
            user_id                = session["user_id"],
            book_title             = book_title,
            chapter_title          = request.form.get("chapter_title", "").strip() or None,
            book_or_chapter        = request.form.get("book_or_chapter", "").strip() or None,
            translated_title       = request.form.get("translated_title", "").strip() or None,
            proceedings_publisher  = request.form.get("proceedings_publisher", "").strip() or None,
            internal_external      = request.form.get("internal_external", "").strip() or None,
            national_international = request.form.get("national_international", "").strip() or None,
            publication_date       = (datetime.strptime(request.form["publication_date"], "%Y-%m-%d").date()
                                      if request.form.get("publication_date") else None),
            isbn                   = request.form.get("isbn", "").strip() or None,
            co_authors             = request.form.get("co_authors", "").strip() or None,
            doi                    = request.form.get("doi", "").strip() or None,
            indexed_in             = request.form.get("indexed_in", "").strip() or None,
            journal_link           = request.form.get("journal_link", "").strip() or None,
            supporting_doc_link    = request.form.get("supporting_doc_link", "").strip() or None,
        )
        db.session.add(entry)
        db.session.commit()
        log_action(f"Faculty added book chapter: {book_title}")
        flash("Book / Chapter added successfully.", "success")
        return redirect(url_for("my_book_chapters"))
    return render_template("faculty/add_book_chapter.html")


# =========================================================
# EDIT BOOK CHAPTER
# templates/faculty/edit_book_chapter.html
# =========================================================

@app.route("/edit_book_chapter/<int:entry_id>", methods=["GET", "POST"])
def edit_book_chapter(entry_id):
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    entry = FacultyBookChapter.query.filter_by(id=entry_id, user_id=session["user_id"]).first_or_404()
    if request.method == "POST":
        book_title = request.form.get("book_title", "").strip()
        if not book_title:
            flash("Title of the Book Published is required.", "error")
            return redirect(url_for("edit_book_chapter", entry_id=entry_id))
        entry.book_title             = book_title
        entry.chapter_title          = request.form.get("chapter_title", "").strip() or None
        entry.book_or_chapter        = request.form.get("book_or_chapter", "").strip() or None
        entry.translated_title       = request.form.get("translated_title", "").strip() or None
        entry.proceedings_publisher  = request.form.get("proceedings_publisher", "").strip() or None
        entry.internal_external      = request.form.get("internal_external", "").strip() or None
        entry.national_international = request.form.get("national_international", "").strip() or None
        entry.publication_date       = (datetime.strptime(request.form["publication_date"], "%Y-%m-%d").date()
                                         if request.form.get("publication_date") else None)
        entry.isbn                   = request.form.get("isbn", "").strip() or None
        entry.co_authors             = request.form.get("co_authors", "").strip() or None
        entry.doi                    = request.form.get("doi", "").strip() or None
        entry.indexed_in             = request.form.get("indexed_in", "").strip() or None
        entry.journal_link           = request.form.get("journal_link", "").strip() or None
        entry.supporting_doc_link    = request.form.get("supporting_doc_link", "").strip() or None
        db.session.commit()
        log_action(f"Faculty edited book chapter #{entry_id}")
        flash("Book / Chapter updated successfully.", "success")
        return redirect(url_for("my_book_chapters"))
    return render_template("faculty/edit_book_chapter.html", entry=entry)


# =========================================================
# DELETE BOOK CHAPTER
# =========================================================

@app.route("/delete_book_chapter/<int:entry_id>", methods=["POST"])
def delete_book_chapter(entry_id):
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    entry = FacultyBookChapter.query.filter_by(id=entry_id, user_id=session["user_id"]).first_or_404()
    db.session.delete(entry)
    db.session.commit()
    log_action(f"Faculty deleted book chapter #{entry_id}")
    flash("Book / Chapter deleted.", "success")
    return redirect(url_for("my_book_chapters"))


# =========================================================
# BULK UPLOAD BOOK CHAPTERS
# templates/faculty/bulk_upload_book_chapters.html
# =========================================================

@app.route("/bulk_upload_book_chapters", methods=["GET", "POST"])
def bulk_upload_book_chapters():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("bulk_upload_book_chapters"))
        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("csv", "xlsx", "xls"):
            flash("Only CSV or Excel (.xlsx/.xls) files are supported.", "error")
            return redirect(url_for("bulk_upload_book_chapters"))
        file_bytes = file.read()
        records, errors = parse_book_chapters(file_bytes, file.filename)
        inserted = 0
        for rec in records:
            entry = FacultyBookChapter(
                user_id                = session["user_id"],
                book_title             = rec["book_title"],
                chapter_title          = rec.get("chapter_title"),
                book_or_chapter        = rec.get("book_or_chapter"),
                translated_title       = rec.get("translated_title"),
                proceedings_publisher  = rec.get("proceedings_publisher"),
                internal_external      = rec.get("internal_external"),
                national_international = rec.get("national_international"),
                publication_date       = rec.get("publication_date"),
                isbn                   = rec.get("isbn"),
                co_authors             = rec.get("co_authors"),
                doi                    = rec.get("doi"),
                indexed_in             = rec.get("indexed_in"),
                journal_link           = rec.get("journal_link"),
                supporting_doc_link    = rec.get("supporting_doc_link"),
            )
            db.session.add(entry)
            inserted += 1
        db.session.commit()
        log_action(f"Faculty bulk-uploaded {inserted} book chapter(s)")
        if errors:
            skipped = ", ".join(f"row {r}: {m}" for r, m in errors[:5])
            flash(f"{inserted} record(s) imported. Skipped rows — {skipped}", "warning")
        else:
            flash(f"{inserted} book chapter(s) imported successfully.", "success")
        return redirect(url_for("my_book_chapters"))
    return render_template("faculty/bulk_upload_book_chapters.html")


# =========================================================
# ADMIN — VIEW ALL BOOK CHAPTERS
# templates/admin/view_book_chapters.html
# =========================================================

@app.route("/admin/view_book_chapters")
def view_book_chapters():
    if session.get("role") != "admin":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()
    query = FacultyBookChapter.query
    if q:
        query = query.filter(
            FacultyBookChapter.book_title.ilike(f"%{q}%") |
            FacultyBookChapter.chapter_title.ilike(f"%{q}%") |
            FacultyBookChapter.co_authors.ilike(f"%{q}%")
        )
    chapters = query.order_by(FacultyBookChapter.publication_date.desc()).all()
    return render_template("admin/view_book_chapters.html", chapters=chapters, q=q)


# =========================================================
# MY GUEST LECTURES
# templates/faculty/my_guest_lectures.html
# =========================================================

@app.route("/my_guest_lectures")
def my_guest_lectures():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()
    query = FacultyGuestLecture.query.filter_by(user_id=session["user_id"])
    if q:
        query = query.filter(
            FacultyGuestLecture.lecture_title.ilike(f"%{q}%") |
            FacultyGuestLecture.organization_location.ilike(f"%{q}%")
        )
    lectures = query.order_by(FacultyGuestLecture.lecture_date.desc()).all()
    return render_template("faculty/my_guest_lectures.html", lectures=lectures, q=q)


# =========================================================
# ADD GUEST LECTURE
# templates/faculty/add_guest_lecture.html
# =========================================================

@app.route("/add_guest_lecture", methods=["GET", "POST"])
def add_guest_lecture():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    if request.method == "POST":
        lecture_title = request.form.get("lecture_title", "").strip()
        if not lecture_title:
            flash("Lecture Title / Topic is required.", "error")
            return redirect(url_for("add_guest_lecture"))
        entry = FacultyGuestLecture(
            user_id               = session["user_id"],
            lecture_title         = lecture_title,
            organization_location = request.form.get("organization_location", "").strip() or None,
            lecture_date          = (datetime.strptime(request.form["lecture_date"], "%Y-%m-%d").date()
                                     if request.form.get("lecture_date") else None),
            jain_or_outside       = request.form.get("jain_or_outside", "").strip() or None,
            mode                  = request.form.get("mode", "").strip() or None,
            audience_type         = request.form.get("audience_type", "").strip() or None,
            brochure_link         = request.form.get("brochure_link", "").strip() or None,
            supporting_doc_link   = request.form.get("supporting_doc_link", "").strip() or None,
        )
        db.session.add(entry)
        db.session.commit()
        log_action(f"Faculty added guest lecture: {lecture_title}")
        flash("Guest Lecture added successfully.", "success")
        return redirect(url_for("my_guest_lectures"))
    return render_template("faculty/add_guest_lecture.html")


# =========================================================
# EDIT GUEST LECTURE
# templates/faculty/edit_guest_lecture.html
# =========================================================

@app.route("/edit_guest_lecture/<int:entry_id>", methods=["GET", "POST"])
def edit_guest_lecture(entry_id):
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    entry = FacultyGuestLecture.query.filter_by(id=entry_id, user_id=session["user_id"]).first_or_404()
    if request.method == "POST":
        lecture_title = request.form.get("lecture_title", "").strip()
        if not lecture_title:
            flash("Lecture Title / Topic is required.", "error")
            return redirect(url_for("edit_guest_lecture", entry_id=entry_id))
        entry.lecture_title         = lecture_title
        entry.organization_location = request.form.get("organization_location", "").strip() or None
        entry.lecture_date          = (datetime.strptime(request.form["lecture_date"], "%Y-%m-%d").date()
                                        if request.form.get("lecture_date") else None)
        entry.jain_or_outside       = request.form.get("jain_or_outside", "").strip() or None
        entry.mode                  = request.form.get("mode", "").strip() or None
        entry.audience_type         = request.form.get("audience_type", "").strip() or None
        entry.brochure_link         = request.form.get("brochure_link", "").strip() or None
        entry.supporting_doc_link   = request.form.get("supporting_doc_link", "").strip() or None
        db.session.commit()
        log_action(f"Faculty edited guest lecture #{entry_id}")
        flash("Guest Lecture updated successfully.", "success")
        return redirect(url_for("my_guest_lectures"))
    return render_template("faculty/edit_guest_lecture.html", entry=entry)


# =========================================================
# DELETE GUEST LECTURE
# =========================================================

@app.route("/delete_guest_lecture/<int:entry_id>", methods=["POST"])
def delete_guest_lecture(entry_id):
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    entry = FacultyGuestLecture.query.filter_by(id=entry_id, user_id=session["user_id"]).first_or_404()
    db.session.delete(entry)
    db.session.commit()
    log_action(f"Faculty deleted guest lecture #{entry_id}")
    flash("Guest Lecture deleted.", "success")
    return redirect(url_for("my_guest_lectures"))


# =========================================================
# BULK UPLOAD GUEST LECTURES
# templates/faculty/bulk_upload_guest_lectures.html
# =========================================================

@app.route("/bulk_upload_guest_lectures", methods=["GET", "POST"])
def bulk_upload_guest_lectures():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("bulk_upload_guest_lectures"))
        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("csv", "xlsx", "xls"):
            flash("Only CSV or Excel (.xlsx/.xls) files are supported.", "error")
            return redirect(url_for("bulk_upload_guest_lectures"))
        file_bytes = file.read()
        records, errors = parse_guest_lectures(file_bytes, file.filename)
        inserted = 0
        for rec in records:
            entry = FacultyGuestLecture(
                user_id               = session["user_id"],
                lecture_title         = rec["lecture_title"],
                organization_location = rec.get("organization_location"),
                lecture_date          = rec.get("lecture_date"),
                jain_or_outside       = rec.get("jain_or_outside"),
                mode                  = rec.get("mode"),
                audience_type         = rec.get("audience_type"),
                brochure_link         = rec.get("brochure_link"),
                supporting_doc_link   = rec.get("supporting_doc_link"),
            )
            db.session.add(entry)
            inserted += 1
        db.session.commit()
        log_action(f"Faculty bulk-uploaded {inserted} guest lecture(s)")
        if errors:
            skipped = ", ".join(f"row {r}: {m}" for r, m in errors[:5])
            flash(f"{inserted} record(s) imported. Skipped rows — {skipped}", "warning")
        else:
            flash(f"{inserted} guest lecture(s) imported successfully.", "success")
        return redirect(url_for("my_guest_lectures"))
    return render_template("faculty/bulk_upload_guest_lectures.html")


# =========================================================
# ADMIN — VIEW ALL GUEST LECTURES
# templates/admin/view_guest_lectures.html
# =========================================================

@app.route("/admin/view_guest_lectures")
def view_guest_lectures():
    if session.get("role") != "admin":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()
    query = FacultyGuestLecture.query
    if q:
        query = query.filter(
            FacultyGuestLecture.lecture_title.ilike(f"%{q}%") |
            FacultyGuestLecture.organization_location.ilike(f"%{q}%")
        )
    lectures = query.order_by(FacultyGuestLecture.lecture_date.desc()).all()
    return render_template("admin/view_guest_lectures.html", lectures=lectures, q=q)


# =========================================================
# MY PATENTS (Faculty)
# templates/faculty/my_patents.html
# =========================================================

@app.route("/my_patents")
def my_patents():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()
    query = FacultyPatent.query.filter_by(user_id=session["user_id"])
    if q:
        query = query.filter(
            FacultyPatent.title.ilike(f"%{q}%") |
            FacultyPatent.application_number.ilike(f"%{q}%") |
            FacultyPatent.awarding_agency.ilike(f"%{q}%")
        )
    patents = query.order_by(FacultyPatent.filing_date.desc()).all()
    return render_template("faculty/my_patents.html", patents=patents, q=q)


# =========================================================
# ADD PATENT
# templates/faculty/add_patent.html
# =========================================================

@app.route("/add_patent", methods=["GET", "POST"])
def add_patent():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        title = request.form.get("title", "").strip()
        if not title:
            flash("Title of Invention / Work is required.", "error")
            return redirect(url_for("add_patent"))

        from datetime import datetime as _dt

        def _parse_date(val):
            val = (val or "").strip()
            if not val:
                return None
            for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d-%b-%Y"):
                try:
                    return _dt.strptime(val, fmt).date()
                except ValueError:
                    continue
            return None

        entry = FacultyPatent(
            user_id=session["user_id"],
            application_number=request.form.get("application_number", "").strip() or None,
            ip_type=request.form.get("ip_type", "").strip() or None,
            title=title,
            status=request.form.get("status", "").strip() or None,
            filing_date=_parse_date(request.form.get("filing_date")),
            published_date=_parse_date(request.form.get("published_date")),
            grant_date=_parse_date(request.form.get("grant_date")),
            awarding_agency=request.form.get("awarding_agency", "").strip() or None,
            national_international=request.form.get("national_international", "").strip() or None,
            commercialization_details=request.form.get("commercialization_details", "").strip() or None,
            oer_contribution=request.form.get("oer_contribution", "").strip() or None,
            supporting_doc_link=request.form.get("supporting_doc_link", "").strip() or None,
        )
        db.session.add(entry)
        db.session.commit()
        log_action(f"Added patent: {title}")
        flash("Patent / IP record added successfully!", "success")
        return redirect(url_for("my_patents"))

    return render_template("faculty/add_patent.html")


# =========================================================
# EDIT PATENT
# templates/faculty/edit_patent.html
# =========================================================

@app.route("/edit_patent/<int:entry_id>", methods=["GET", "POST"])
def edit_patent(entry_id):
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    entry = FacultyPatent.query.get_or_404(entry_id)
    if entry.user_id != session["user_id"]:
        abort(403)

    if request.method == "POST":
        title = request.form.get("title", "").strip()
        if not title:
            flash("Title of Invention / Work is required.", "error")
            return redirect(url_for("edit_patent", entry_id=entry_id))

        from datetime import datetime as _dt

        def _parse_date(val):
            val = (val or "").strip()
            if not val:
                return None
            for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d-%b-%Y"):
                try:
                    return _dt.strptime(val, fmt).date()
                except ValueError:
                    continue
            return None

        entry.application_number = request.form.get("application_number", "").strip() or None
        entry.ip_type = request.form.get("ip_type", "").strip() or None
        entry.title = title
        entry.status = request.form.get("status", "").strip() or None
        entry.filing_date = _parse_date(request.form.get("filing_date"))
        entry.published_date = _parse_date(request.form.get("published_date"))
        entry.grant_date = _parse_date(request.form.get("grant_date"))
        entry.awarding_agency = request.form.get("awarding_agency", "").strip() or None
        entry.national_international = request.form.get("national_international", "").strip() or None
        entry.commercialization_details = request.form.get("commercialization_details", "").strip() or None
        entry.oer_contribution = request.form.get("oer_contribution", "").strip() or None
        entry.supporting_doc_link = request.form.get("supporting_doc_link", "").strip() or None

        db.session.commit()
        log_action(f"Edited patent #{entry_id}: {title}")
        flash("Patent / IP record updated.", "success")
        return redirect(url_for("my_patents"))

    return render_template("faculty/edit_patent.html", entry=entry)


# =========================================================
# DELETE PATENT
# =========================================================

@app.route("/delete_patent/<int:entry_id>", methods=["POST"])
def delete_patent(entry_id):
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    entry = FacultyPatent.query.get_or_404(entry_id)
    if entry.user_id != session["user_id"]:
        abort(403)
    db.session.delete(entry)
    db.session.commit()
    log_action(f"Deleted patent #{entry_id}: {entry.title}")
    flash("Patent / IP record deleted.", "success")
    return redirect(url_for("my_patents"))


# =========================================================
# BULK UPLOAD PATENTS
# templates/faculty/bulk_upload_patents.html
# =========================================================

@app.route("/bulk_upload_patents", methods=["GET", "POST"])
def bulk_upload_patents():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file.", "error")
            return redirect(url_for("bulk_upload_patents"))

        file_bytes = file.read()
        records, errors = parse_patents(file_bytes, file.filename)
        added = 0
        for rec in records:
            entry = FacultyPatent(user_id=session["user_id"], **rec)
            db.session.add(entry)
            added += 1
        db.session.commit()
        log_action(f"Bulk uploaded {added} patent(s)")
        if errors:
            for row_num, msg in errors:
                flash(f"Row {row_num}: {msg}", "error")
        flash(f"Successfully imported {added} patent record(s).", "success")
        return redirect(url_for("my_patents"))

    return render_template("faculty/bulk_upload_patents.html")


# =========================================================
# ADMIN — VIEW ALL PATENTS
# templates/admin/view_patents.html
# =========================================================

@app.route("/admin/view_patents")
def view_patents():
    if session.get("role") != "admin":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()
    query = FacultyPatent.query
    if q:
        query = query.filter(
            FacultyPatent.title.ilike(f"%{q}%") |
            FacultyPatent.application_number.ilike(f"%{q}%") |
            FacultyPatent.awarding_agency.ilike(f"%{q}%")
        )
    patents = query.order_by(FacultyPatent.filing_date.desc()).all()
    return render_template("admin/view_patents.html", patents=patents, q=q)


# =========================================================
# MY FELLOWSHIPS (Faculty)
# templates/faculty/my_fellowships.html
# =========================================================

@app.route("/my_fellowships")
def my_fellowships():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()
    query = FacultyFellowship.query.filter_by(user_id=session["user_id"])
    if q:
        query = query.filter(
            FacultyFellowship.award_name.ilike(f"%{q}%") |
            FacultyFellowship.awarding_agency.ilike(f"%{q}%") |
            FacultyFellowship.research_topic.ilike(f"%{q}%")
        )
    fellowships = query.order_by(FacultyFellowship.award_date.desc()).all()
    return render_template("faculty/my_fellowships.html", fellowships=fellowships, q=q)


# =========================================================
# ADD FELLOWSHIP
# templates/faculty/add_fellowship.html
# =========================================================

@app.route("/add_fellowship", methods=["GET", "POST"])
def add_fellowship():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        award_name = request.form.get("award_name", "").strip()
        if not award_name:
            flash("Name of the award / fellowship is required.", "error")
            return redirect(url_for("add_fellowship"))

        from datetime import datetime as _dt

        def _parse_date(val):
            val = (val or "").strip()
            if not val:
                return None
            for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d-%b-%Y"):
                try:
                    return _dt.strptime(val, fmt).date()
                except ValueError:
                    continue
            return None

        entry = FacultyFellowship(
            user_id=session["user_id"],
            award_name=award_name,
            financial_support=request.form.get("financial_support", "").strip() or None,
            grant_purpose=request.form.get("grant_purpose", "").strip() or None,
            support_type=request.form.get("support_type", "").strip() or None,
            national_international=request.form.get("national_international", "").strip() or None,
            award_date=_parse_date(request.form.get("award_date")),
            awarding_agency=request.form.get("awarding_agency", "").strip() or None,
            duration=request.form.get("duration", "").strip() or None,
            research_topic=request.form.get("research_topic", "").strip() or None,
            location=request.form.get("location", "").strip() or None,
            collaborating_institution=request.form.get("collaborating_institution", "").strip() or None,
            grant_letter_link=request.form.get("grant_letter_link", "").strip() or None,
        )
        db.session.add(entry)
        db.session.commit()
        log_action(f"Added fellowship: {award_name}")
        flash("Fellowship record added successfully!", "success")
        return redirect(url_for("my_fellowships"))

    return render_template("faculty/add_fellowship.html")


# =========================================================
# EDIT FELLOWSHIP
# templates/faculty/edit_fellowship.html
# =========================================================

@app.route("/edit_fellowship/<int:entry_id>", methods=["GET", "POST"])
def edit_fellowship(entry_id):
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    entry = FacultyFellowship.query.get_or_404(entry_id)
    if entry.user_id != session["user_id"]:
        abort(403)

    if request.method == "POST":
        award_name = request.form.get("award_name", "").strip()
        if not award_name:
            flash("Name of the award / fellowship is required.", "error")
            return redirect(url_for("edit_fellowship", entry_id=entry_id))

        from datetime import datetime as _dt

        def _parse_date(val):
            val = (val or "").strip()
            if not val:
                return None
            for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d-%b-%Y"):
                try:
                    return _dt.strptime(val, fmt).date()
                except ValueError:
                    continue
            return None

        entry.award_name = award_name
        entry.financial_support = request.form.get("financial_support", "").strip() or None
        entry.grant_purpose = request.form.get("grant_purpose", "").strip() or None
        entry.support_type = request.form.get("support_type", "").strip() or None
        entry.national_international = request.form.get("national_international", "").strip() or None
        entry.award_date = _parse_date(request.form.get("award_date"))
        entry.awarding_agency = request.form.get("awarding_agency", "").strip() or None
        entry.duration = request.form.get("duration", "").strip() or None
        entry.research_topic = request.form.get("research_topic", "").strip() or None
        entry.location = request.form.get("location", "").strip() or None
        entry.collaborating_institution = request.form.get("collaborating_institution", "").strip() or None
        entry.grant_letter_link = request.form.get("grant_letter_link", "").strip() or None

        db.session.commit()
        log_action(f"Edited fellowship #{entry_id}: {award_name}")
        flash("Fellowship record updated.", "success")
        return redirect(url_for("my_fellowships"))

    return render_template("faculty/edit_fellowship.html", entry=entry)


# =========================================================
# DELETE FELLOWSHIP
# =========================================================

@app.route("/delete_fellowship/<int:entry_id>", methods=["POST"])
def delete_fellowship(entry_id):
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    entry = FacultyFellowship.query.get_or_404(entry_id)
    if entry.user_id != session["user_id"]:
        abort(403)
    db.session.delete(entry)
    db.session.commit()
    log_action(f"Deleted fellowship #{entry_id}: {entry.award_name}")
    flash("Fellowship record deleted.", "success")
    return redirect(url_for("my_fellowships"))


# =========================================================
# BULK UPLOAD FELLOWSHIPS
# templates/faculty/bulk_upload_fellowships.html
# =========================================================

@app.route("/bulk_upload_fellowships", methods=["GET", "POST"])
def bulk_upload_fellowships():
    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file.", "error")
            return redirect(url_for("bulk_upload_fellowships"))

        file_bytes = file.read()
        records, errors = parse_fellowships(file_bytes, file.filename)
        added = 0
        for rec in records:
            entry = FacultyFellowship(user_id=session["user_id"], **rec)
            db.session.add(entry)
            added += 1
        db.session.commit()
        log_action(f"Bulk uploaded {added} fellowship(s)")
        if errors:
            for row_num, msg in errors:
                flash(f"Row {row_num}: {msg}", "error")
        flash(f"Successfully imported {added} fellowship record(s).", "success")
        return redirect(url_for("my_fellowships"))

    return render_template("faculty/bulk_upload_fellowships.html")


# =========================================================
# ADMIN — VIEW ALL FELLOWSHIPS
# templates/admin/view_fellowships.html
# =========================================================

@app.route("/admin/view_fellowships")
def view_fellowships():
    if session.get("role") != "admin":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()
    query = FacultyFellowship.query
    if q:
        query = query.filter(
            FacultyFellowship.award_name.ilike(f"%{q}%") |
            FacultyFellowship.awarding_agency.ilike(f"%{q}%") |
            FacultyFellowship.research_topic.ilike(f"%{q}%")
        )
    fellowships = query.order_by(FacultyFellowship.award_date.desc()).all()
    return render_template("admin/view_fellowships.html", fellowships=fellowships, q=q)


# =========================================================
# CONFERENCES — COMBINED LIST (participated + organised)
# templates/faculty/my_conferences.html
# =========================================================

@app.route("/my_conferences")
def my_conferences():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    participated = ConferenceParticipated.query.filter_by(
        user_id=session["user_id"]
    ).order_by(ConferenceParticipated.created_at.desc()).all()

    organised = ConferenceOrganised.query.filter_by(
        user_id=session["user_id"]
    ).order_by(ConferenceOrganised.created_at.desc()).all()

    log_action("Faculty viewed Conferences")

    return render_template(
        "faculty/my_conferences.html",
        participated=participated,
        organised=organised
    )


# =========================================================
# CONFERENCES — ADD PARTICIPATED
# =========================================================

@app.route("/add_conference_participated", methods=["GET", "POST"])
def add_conference_participated():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":

        conference_title    = request.form.get("conference_title")
        paper_title         = request.form.get("paper_title")
        organisers          = request.form.get("organisers")
        leads_to_pub        = request.form.get("leads_to_publication")
        paper_link          = request.form.get("paper_link")
        collaboration       = request.form.get("collaboration")
        focus_area          = request.form.get("focus_area")
        date_from_raw       = request.form.get("date_from")
        date_to_raw         = request.form.get("date_to")
        nat_int             = request.form.get("national_international")
        proceedings_indexed = request.form.get("proceedings_indexed")
        indexing_details    = request.form.get("indexing_details")
        cert_link           = request.form.get("certificate_link")

        entry = ConferenceParticipated(
            user_id=session["user_id"],
            conference_title=conference_title,
            paper_title=paper_title,
            organisers=organisers,
            leads_to_publication=leads_to_pub,
            paper_link=paper_link,
            collaboration=collaboration,
            focus_area=focus_area,
            national_international=nat_int,
            proceedings_indexed=proceedings_indexed,
            indexing_details=indexing_details,
            certificate_link=cert_link
        )

        if date_from_raw:
            try:
                entry.date_from = datetime.strptime(date_from_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        if date_to_raw:
            try:
                entry.date_to = datetime.strptime(date_to_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        db.session.add(entry)
        db.session.commit()
        log_action(f"Faculty added conference participated: {conference_title}")
        flash("Conference participated added successfully.", "success")
        return redirect(url_for("my_conferences"))

    return render_template("faculty/add_conference_participated.html")


# =========================================================
# CONFERENCES — EDIT PARTICIPATED
# =========================================================

@app.route("/edit_conference_participated/<int:entry_id>", methods=["GET", "POST"])
def edit_conference_participated(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = ConferenceParticipated.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    if request.method == "POST":

        entry.conference_title    = request.form.get("conference_title")
        entry.paper_title         = request.form.get("paper_title")
        entry.organisers          = request.form.get("organisers")
        entry.leads_to_publication = request.form.get("leads_to_publication")
        entry.paper_link          = request.form.get("paper_link")
        entry.collaboration       = request.form.get("collaboration")
        entry.focus_area          = request.form.get("focus_area")
        entry.national_international = request.form.get("national_international")
        entry.proceedings_indexed = request.form.get("proceedings_indexed")
        entry.indexing_details    = request.form.get("indexing_details")
        entry.certificate_link    = request.form.get("certificate_link")

        date_from_raw = request.form.get("date_from")
        date_to_raw   = request.form.get("date_to")
        if date_from_raw:
            try:
                entry.date_from = datetime.strptime(date_from_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        if date_to_raw:
            try:
                entry.date_to = datetime.strptime(date_to_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        db.session.commit()
        log_action(f"Faculty edited conference participated: {entry.conference_title}")
        flash("Conference updated.", "success")
        return redirect(url_for("my_conferences"))

    return render_template("faculty/edit_conference_participated.html", entry=entry)


# =========================================================
# CONFERENCES — DELETE PARTICIPATED
# =========================================================

@app.route("/delete_conference_participated/<int:entry_id>", methods=["POST"])
def delete_conference_participated(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = ConferenceParticipated.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    db.session.delete(entry)
    db.session.commit()
    log_action(f"Faculty deleted conference participated: {entry.conference_title}")
    flash("Entry deleted.", "success")
    return redirect(url_for("my_conferences"))


# =========================================================
# CONFERENCES — ADD ORGANISED
# =========================================================

@app.route("/add_conference_organised", methods=["GET", "POST"])
def add_conference_organised():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":

        title               = request.form.get("title")
        department          = request.form.get("department")
        faculty_role        = request.form.get("faculty_role")
        collaboration       = request.form.get("collaboration")
        focus_area          = request.form.get("focus_area")
        nat_int             = request.form.get("national_international")
        num_participants    = request.form.get("num_participants") or None
        date_from_raw       = request.form.get("date_from")
        date_to_raw         = request.form.get("date_to")
        proceedings_indexed = request.form.get("proceedings_indexed")
        indexing_details    = request.form.get("indexing_details")
        report_link         = request.form.get("activity_report_link")

        entry = ConferenceOrganised(
            user_id=session["user_id"],
            title=title,
            department=department,
            faculty_role=faculty_role,
            collaboration=collaboration,
            focus_area=focus_area,
            national_international=nat_int,
            num_participants=int(num_participants) if num_participants else None,
            proceedings_indexed=proceedings_indexed,
            indexing_details=indexing_details,
            activity_report_link=report_link
        )

        if date_from_raw:
            try:
                entry.date_from = datetime.strptime(date_from_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        if date_to_raw:
            try:
                entry.date_to = datetime.strptime(date_to_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        db.session.add(entry)
        db.session.commit()
        log_action(f"Faculty added conference organised: {title}")
        flash("Conference organised added successfully.", "success")
        return redirect(url_for("my_conferences"))

    return render_template("faculty/add_conference_organised.html")


# =========================================================
# CONFERENCES — EDIT ORGANISED
# =========================================================

@app.route("/edit_conference_organised/<int:entry_id>", methods=["GET", "POST"])
def edit_conference_organised(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = ConferenceOrganised.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    if request.method == "POST":

        entry.title               = request.form.get("title")
        entry.department          = request.form.get("department")
        entry.faculty_role        = request.form.get("faculty_role")
        entry.collaboration       = request.form.get("collaboration")
        entry.focus_area          = request.form.get("focus_area")
        entry.national_international = request.form.get("national_international")
        np = request.form.get("num_participants")
        entry.num_participants    = int(np) if np else None
        entry.proceedings_indexed = request.form.get("proceedings_indexed")
        entry.indexing_details    = request.form.get("indexing_details")
        entry.activity_report_link = request.form.get("activity_report_link")

        date_from_raw = request.form.get("date_from")
        date_to_raw   = request.form.get("date_to")
        if date_from_raw:
            try:
                entry.date_from = datetime.strptime(date_from_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        if date_to_raw:
            try:
                entry.date_to = datetime.strptime(date_to_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        db.session.commit()
        log_action(f"Faculty edited conference organised: {entry.title}")
        flash("Conference updated.", "success")
        return redirect(url_for("my_conferences"))

    return render_template("faculty/edit_conference_organised.html", entry=entry)


# =========================================================
# CONFERENCES — DELETE ORGANISED
# =========================================================

@app.route("/delete_conference_organised/<int:entry_id>", methods=["POST"])
def delete_conference_organised(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = ConferenceOrganised.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    db.session.delete(entry)
    db.session.commit()
    log_action(f"Faculty deleted conference organised: {entry.title}")
    flash("Entry deleted.", "success")
    return redirect(url_for("my_conferences"))


# =========================================================
# CONFERENCES — BULK UPLOAD PARTICIPATED
# =========================================================

@app.route("/bulk_upload_conferences_participated", methods=["GET", "POST"])
def bulk_upload_conferences_participated():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("bulk_upload_conferences_participated"))

        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("csv", "xlsx", "xls"):
            flash("Only CSV or Excel (.xlsx/.xls) files are supported.", "error")
            return redirect(url_for("bulk_upload_conferences_participated"))

        file_bytes = file.read()
        records, errors = parse_conferences_participated(file_bytes, file.filename)

        inserted = 0
        for rec in records:
            entry = ConferenceParticipated(
                user_id              = session["user_id"],
                conference_title     = rec["conference_title"],
                paper_title          = rec["paper_title"],
                organisers           = rec["organisers"],
                leads_to_publication = rec["leads_to_publication"],
                paper_link           = rec["paper_link"],
                collaboration        = rec["collaboration"],
                focus_area           = rec["focus_area"],
                date_from            = rec["date_from"],
                date_to              = rec["date_to"],
                national_international = rec["national_international"],
                proceedings_indexed  = rec["proceedings_indexed"],
                indexing_details     = rec["indexing_details"],
                certificate_link     = rec["certificate_link"],
            )
            db.session.add(entry)
            inserted += 1

        db.session.commit()
        log_action(f"Faculty bulk-uploaded {inserted} conference(s) participated")

        if errors:
            skipped = ", ".join(f"row {r}: {m}" for r, m in errors[:5])
            flash(f"{inserted} record(s) imported. Skipped rows — {skipped}", "warning")
        else:
            flash(f"{inserted} conference(s) participated imported successfully.", "success")

        return redirect(url_for("my_conferences"))

    return render_template("faculty/bulk_upload_conferences_participated.html")


# =========================================================
# CONFERENCES — BULK UPLOAD ORGANISED
# =========================================================

@app.route("/bulk_upload_conferences_organised", methods=["GET", "POST"])
def bulk_upload_conferences_organised():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("bulk_upload_conferences_organised"))

        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("csv", "xlsx", "xls"):
            flash("Only CSV or Excel (.xlsx/.xls) files are supported.", "error")
            return redirect(url_for("bulk_upload_conferences_organised"))

        file_bytes = file.read()
        records, errors = parse_conferences_organised(file_bytes, file.filename)

        inserted = 0
        for rec in records:
            entry = ConferenceOrganised(
                user_id              = session["user_id"],
                title                = rec["title"],
                department           = rec["department"],
                faculty_role         = rec["faculty_role"],
                collaboration        = rec["collaboration"],
                focus_area           = rec["focus_area"],
                national_international = rec["national_international"],
                num_participants     = rec["num_participants"],
                date_from            = rec["date_from"],
                date_to              = rec["date_to"],
                proceedings_indexed  = rec["proceedings_indexed"],
                indexing_details     = rec["indexing_details"],
                activity_report_link = rec["activity_report_link"],
            )
            db.session.add(entry)
            inserted += 1

        db.session.commit()
        log_action(f"Faculty bulk-uploaded {inserted} conference(s) organised")

        if errors:
            skipped = ", ".join(f"row {r}: {m}" for r, m in errors[:5])
            flash(f"{inserted} record(s) imported. Skipped rows — {skipped}", "warning")
        else:
            flash(f"{inserted} conference(s) organised imported successfully.", "success")

        return redirect(url_for("my_conferences"))

    return render_template("faculty/bulk_upload_conferences_organised.html")


# =========================================================
# ADMIN — VIEW CONFERENCES
# templates/admin/view_conferences.html
# =========================================================

@app.route("/admin/view_conferences")
def view_conferences():
    if session.get("role") != "admin":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()

    p_query = ConferenceParticipated.query
    o_query = ConferenceOrganised.query
    if q:
        p_query = p_query.filter(
            ConferenceParticipated.conference_title.ilike(f"%{q}%") |
            ConferenceParticipated.organisers.ilike(f"%{q}%") |
            ConferenceParticipated.paper_title.ilike(f"%{q}%")
        )
        o_query = o_query.filter(
            ConferenceOrganised.title.ilike(f"%{q}%") |
            ConferenceOrganised.department.ilike(f"%{q}%") |
            ConferenceOrganised.faculty_role.ilike(f"%{q}%")
        )
    participated = p_query.order_by(ConferenceParticipated.date_from.desc()).all()
    organised = o_query.order_by(ConferenceOrganised.date_from.desc()).all()
    return render_template(
        "admin/view_conferences.html",
        participated=participated,
        organised=organised,
        q=q
    )


# =========================================================
# FDP PROGRAMS — MY LIST (Combined)
# templates/faculty/my_fdp_programs.html
# =========================================================

@app.route("/my_fdp_programs")
def my_fdp_programs():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    participated = FDPParticipated.query.filter_by(
        user_id=session["user_id"]
    ).order_by(FDPParticipated.created_at.desc()).all()

    organised = FDPOrganised.query.filter_by(
        user_id=session["user_id"]
    ).order_by(FDPOrganised.created_at.desc()).all()

    log_action("Faculty viewed FDP Programs")

    return render_template(
        "faculty/my_fdp_programs.html",
        participated=participated,
        organised=organised
    )


# =========================================================
# FDP — ADD PARTICIPATED
# =========================================================

@app.route("/add_fdp_participated", methods=["GET", "POST"])
def add_fdp_participated():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":

        program_title       = request.form.get("program_title")
        duration_days       = request.form.get("duration_days")
        start_date_raw      = request.form.get("start_date")
        end_date_raw        = request.form.get("end_date")
        program_type        = request.form.get("program_type")
        nat_int             = request.form.get("national_international")
        organizing_agency   = request.form.get("organizing_agency")
        location            = request.form.get("location")
        mode                = request.form.get("mode")
        funding             = request.form.get("funding")
        certificate_link    = request.form.get("certificate_link")
        brochure_link       = request.form.get("brochure_link")
        enrolled_coursera   = request.form.get("enrolled_coursera")

        entry = FDPParticipated(
            user_id=session["user_id"],
            program_title=program_title,
            duration_days=duration_days,
            program_type=program_type,
            national_international=nat_int,
            organizing_agency=organizing_agency,
            location=location,
            mode=mode,
            funding=funding,
            certificate_link=certificate_link,
            brochure_link=brochure_link,
            enrolled_coursera=enrolled_coursera
        )

        if start_date_raw:
            try:
                entry.start_date = datetime.strptime(start_date_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        if end_date_raw:
            try:
                entry.end_date = datetime.strptime(end_date_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        db.session.add(entry)
        db.session.commit()
        log_action(f"Faculty added FDP participated: {program_title}")
        flash("FDP participated added successfully.", "success")
        return redirect(url_for("my_fdp_programs"))

    return render_template("faculty/add_fdp_participated.html")


# =========================================================
# FDP — EDIT PARTICIPATED
# =========================================================

@app.route("/edit_fdp_participated/<int:entry_id>", methods=["GET", "POST"])
def edit_fdp_participated(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = FDPParticipated.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    if request.method == "POST":

        entry.program_title       = request.form.get("program_title")
        entry.duration_days       = request.form.get("duration_days")
        entry.program_type        = request.form.get("program_type")
        entry.national_international = request.form.get("national_international")
        entry.organizing_agency   = request.form.get("organizing_agency")
        entry.location            = request.form.get("location")
        entry.mode                = request.form.get("mode")
        entry.funding             = request.form.get("funding")
        entry.certificate_link    = request.form.get("certificate_link")
        entry.brochure_link       = request.form.get("brochure_link")
        entry.enrolled_coursera   = request.form.get("enrolled_coursera")

        start_date_raw = request.form.get("start_date")
        end_date_raw   = request.form.get("end_date")
        if start_date_raw:
            try:
                entry.start_date = datetime.strptime(start_date_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        else:
            entry.start_date = None
        if end_date_raw:
            try:
                entry.end_date = datetime.strptime(end_date_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        else:
            entry.end_date = None

        db.session.commit()
        log_action(f"Faculty edited FDP participated: {entry.program_title}")
        flash("FDP participation updated.", "success")
        return redirect(url_for("my_fdp_programs"))

    return render_template("faculty/edit_fdp_participated.html", entry=entry)


# =========================================================
# FDP — DELETE PARTICIPATED
# =========================================================

@app.route("/delete_fdp_participated/<int:entry_id>", methods=["POST"])
def delete_fdp_participated(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = FDPParticipated.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    db.session.delete(entry)
    db.session.commit()
    log_action(f"Faculty deleted FDP participated: {entry.program_title}")
    flash("Entry deleted.", "success")
    return redirect(url_for("my_fdp_programs"))


# =========================================================
# FDP — ADD ORGANISED
# =========================================================

@app.route("/add_fdp_organised", methods=["GET", "POST"])
def add_fdp_organised():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":

        title               = request.form.get("title")
        department          = request.form.get("department")
        faculty_role        = request.form.get("faculty_role")
        collaboration       = request.form.get("collaboration")
        focus_area          = request.form.get("focus_area")
        nat_int             = request.form.get("national_international")
        num_participants    = request.form.get("num_participants") or None
        date_from_raw       = request.form.get("date_from")
        date_to_raw         = request.form.get("date_to")
        report_link         = request.form.get("activity_report_link")

        entry = FDPOrganised(
            user_id=session["user_id"],
            title=title,
            department=department,
            faculty_role=faculty_role,
            collaboration=collaboration,
            focus_area=focus_area,
            national_international=nat_int,
            num_participants=int(num_participants) if num_participants else None,
            activity_report_link=report_link
        )

        if date_from_raw:
            try:
                entry.date_from = datetime.strptime(date_from_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        if date_to_raw:
            try:
                entry.date_to = datetime.strptime(date_to_raw, "%Y-%m-%d").date()
            except ValueError:
                pass

        db.session.add(entry)
        db.session.commit()
        log_action(f"Faculty added FDP organised: {title}")
        flash("FDP organised added successfully.", "success")
        return redirect(url_for("my_fdp_programs"))

    return render_template("faculty/add_fdp_organised.html")


# =========================================================
# FDP — EDIT ORGANISED
# =========================================================

@app.route("/edit_fdp_organised/<int:entry_id>", methods=["GET", "POST"])
def edit_fdp_organised(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = FDPOrganised.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    if request.method == "POST":

        entry.title               = request.form.get("title")
        entry.department          = request.form.get("department")
        entry.faculty_role        = request.form.get("faculty_role")
        entry.collaboration       = request.form.get("collaboration")
        entry.focus_area          = request.form.get("focus_area")
        entry.national_international = request.form.get("national_international")
        np_val = request.form.get("num_participants")
        entry.num_participants    = int(np_val) if np_val else None
        entry.activity_report_link = request.form.get("activity_report_link")

        date_from_raw = request.form.get("date_from")
        date_to_raw   = request.form.get("date_to")
        if date_from_raw:
            try:
                entry.date_from = datetime.strptime(date_from_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        else:
            entry.date_from = None
        if date_to_raw:
            try:
                entry.date_to = datetime.strptime(date_to_raw, "%Y-%m-%d").date()
            except ValueError:
                pass
        else:
            entry.date_to = None

        db.session.commit()
        log_action(f"Faculty edited FDP organised: {entry.title}")
        flash("FDP organised updated.", "success")
        return redirect(url_for("my_fdp_programs"))

    return render_template("faculty/edit_fdp_organised.html", entry=entry)


# =========================================================
# FDP — DELETE ORGANISED
# =========================================================

@app.route("/delete_fdp_organised/<int:entry_id>", methods=["POST"])
def delete_fdp_organised(entry_id):

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    entry = FDPOrganised.query.filter_by(
        id=entry_id, user_id=session["user_id"]
    ).first_or_404()

    db.session.delete(entry)
    db.session.commit()
    log_action(f"Faculty deleted FDP organised: {entry.title}")
    flash("Entry deleted.", "success")
    return redirect(url_for("my_fdp_programs"))


# =========================================================
# FDP — BULK UPLOAD PARTICIPATED
# =========================================================

@app.route("/bulk_upload_fdp_participated", methods=["GET", "POST"])
def bulk_upload_fdp_participated():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("bulk_upload_fdp_participated"))

        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("csv", "xlsx", "xls"):
            flash("Only CSV or Excel (.xlsx/.xls) files are supported.", "error")
            return redirect(url_for("bulk_upload_fdp_participated"))

        file_bytes = file.read()
        records, errors = parse_fdp_participated(file_bytes, file.filename)

        inserted = 0
        for rec in records:
            entry = FDPParticipated(
                user_id              = session["user_id"],
                program_title        = rec["program_title"],
                duration_days        = rec["duration_days"],
                start_date           = rec["start_date"],
                end_date             = rec["end_date"],
                program_type         = rec["program_type"],
                national_international = rec["national_international"],
                organizing_agency    = rec["organizing_agency"],
                location             = rec["location"],
                mode                 = rec["mode"],
                funding              = rec["funding"],
                certificate_link     = rec["certificate_link"],
                brochure_link        = rec["brochure_link"],
                enrolled_coursera    = rec["enrolled_coursera"],
            )
            db.session.add(entry)
            inserted += 1

        db.session.commit()
        log_action(f"Faculty bulk-uploaded {inserted} FDP(s) participated")

        if errors:
            skipped = ", ".join(f"row {r}: {m}" for r, m in errors[:5])
            flash(f"{inserted} record(s) imported. Skipped rows — {skipped}", "warning")
        else:
            flash(f"{inserted} FDP(s) participated imported successfully.", "success")

        return redirect(url_for("my_fdp_programs"))

    return render_template("faculty/bulk_upload_fdp_participated.html")


# =========================================================
# FDP — BULK UPLOAD ORGANISED
# =========================================================

@app.route("/bulk_upload_fdp_organised", methods=["GET", "POST"])
def bulk_upload_fdp_organised():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Please select a file to upload.", "error")
            return redirect(url_for("bulk_upload_fdp_organised"))

        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("csv", "xlsx", "xls"):
            flash("Only CSV or Excel (.xlsx/.xls) files are supported.", "error")
            return redirect(url_for("bulk_upload_fdp_organised"))

        file_bytes = file.read()
        records, errors = parse_fdp_organised(file_bytes, file.filename)

        inserted = 0
        for rec in records:
            entry = FDPOrganised(
                user_id              = session["user_id"],
                title                = rec["title"],
                department           = rec["department"],
                faculty_role         = rec["faculty_role"],
                collaboration        = rec["collaboration"],
                focus_area           = rec["focus_area"],
                national_international = rec["national_international"],
                num_participants     = rec["num_participants"],
                date_from            = rec["date_from"],
                date_to              = rec["date_to"],
                activity_report_link = rec["activity_report_link"],
            )
            db.session.add(entry)
            inserted += 1

        db.session.commit()
        log_action(f"Faculty bulk-uploaded {inserted} FDP(s) organised")

        if errors:
            skipped = ", ".join(f"row {r}: {m}" for r, m in errors[:5])
            flash(f"{inserted} record(s) imported. Skipped rows — {skipped}", "warning")
        else:
            flash(f"{inserted} FDP(s) organised imported successfully.", "success")

        return redirect(url_for("my_fdp_programs"))

    return render_template("faculty/bulk_upload_fdp_organised.html")


# =========================================================
# ADMIN — VIEW FDP PROGRAMS
# templates/admin/view_fdp_programs.html
# =========================================================

@app.route("/admin/view_fdp_programs")
def view_fdp_programs():
    if session.get("role") != "admin":
        return redirect(url_for("login"))
    q = request.args.get("q", "").strip()

    p_query = FDPParticipated.query
    o_query = FDPOrganised.query
    if q:
        p_query = p_query.filter(
            FDPParticipated.program_title.ilike(f"%{q}%") |
            FDPParticipated.organizing_agency.ilike(f"%{q}%") |
            FDPParticipated.location.ilike(f"%{q}%")
        )
        o_query = o_query.filter(
            FDPOrganised.title.ilike(f"%{q}%") |
            FDPOrganised.department.ilike(f"%{q}%") |
            FDPOrganised.faculty_role.ilike(f"%{q}%")
        )
    participated = p_query.order_by(FDPParticipated.start_date.desc()).all()
    organised = o_query.order_by(FDPOrganised.date_from.desc()).all()
    return render_template(
        "admin/view_fdp_programs.html",
        participated=participated,
        organised=organised,
        q=q
    )


# =========================================================
# CHANGE PASSWORD (Faculty — OTP verified)
# templates/faculty/change_password.html
# =========================================================

@app.route("/change_password", methods=["GET", "POST"])
@limiter.limit("10 per minute")
def change_password():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    user = User.query.get(session["user_id"])
    if not user:
        session.clear()
        return redirect(url_for("login"))

    # Step 1: Faculty lands on GET — send OTP automatically
    if request.method == "GET":
        otp = generate_otp()
        store_otp(user, otp, otp_type="reset")
        send_otp(user, otp)
        log_action("Faculty requested change-password OTP")
        flash("An OTP has been sent to your registered email.", "success")
        return render_template("faculty/change_password.html")

    # Step 2: Faculty submits OTP + new password
    entered_otp      = request.form.get("otp", "").strip()
    new_password      = request.form.get("new_password", "")
    confirm_password  = request.form.get("confirm_password", "")

    if not entered_otp:
        flash("Please enter the OTP.", "error")
        return render_template("faculty/change_password.html")

    if len(new_password) < 8:
        flash("Password must be at least 8 characters.", "error")
        return render_template("faculty/change_password.html")

    if new_password != confirm_password:
        flash("Passwords do not match.", "error")
        return render_template("faculty/change_password.html")

    if not verify_otp(user, entered_otp, otp_type="reset"):
        flash("Invalid or expired OTP.", "error")
        return render_template("faculty/change_password.html")

    # OTP valid — update password
    user.password_hash   = hash_password(new_password)
    user.reset_otp_hash  = None
    user.reset_otp_expiry = None
    db.session.commit()

    log_action("Faculty changed password via OTP")
    flash("Password changed successfully.", "success")
    return redirect(url_for("faculty_dashboard"))


@app.route("/resend_change_password_otp", methods=["POST"])
@limiter.limit("3 per minute")
def resend_change_password_otp():

    if session.get("role") != "faculty":
        return redirect(url_for("login"))

    user = User.query.get(session["user_id"])
    if not user:
        session.clear()
        return redirect(url_for("login"))

    otp = generate_otp()
    store_otp(user, otp, otp_type="reset")
    send_otp(user, otp)
    log_action("Faculty resent change-password OTP")
    flash("A new OTP has been sent to your email.", "success")
    return redirect(url_for("change_password_form"))


@app.route("/change_password_form")
def change_password_form():
    """Render the change-password page without re-sending OTP."""
    if session.get("role") != "faculty":
        return redirect(url_for("login"))
    return render_template("faculty/change_password.html")


# =========================================================
# FORGOT PASSWORD
# templates/auth/forgot_password.html
# =========================================================

@app.route("/forgot_password", methods=["GET", "POST"])
@limiter.limit("5 per minute")
def forgot_password():

    if request.method == "POST":

        email = request.form.get("email", "").strip().lower()

        user = User.query.filter_by(email=email).first()

        # Always show same message to prevent enumeration
        flash("If that email exists, a reset OTP has been sent.", "success")

        if user and user.is_active_account:

            otp = generate_otp()
            store_otp(user, otp, otp_type="reset")
            send_otp(user, otp)

            log_action(f"Password reset OTP sent: {email}")

            session["reset_user_id"] = user.id

        return redirect(url_for("reset_password"))

    return render_template("auth/forgot_password.html")


# =========================================================
# RESET PASSWORD
# templates/auth/reset_password.html
# =========================================================

@app.route("/reset_password", methods=["GET", "POST"])
@limiter.limit("10 per minute")
def reset_password():

    if request.method == "POST":

        user_id = session.get("reset_user_id")

        if not user_id:
            flash("Session expired. Please request a new OTP.", "error")
            return redirect(url_for("forgot_password"))

        user = User.query.get(user_id)

        if not user:
            session.clear()
            flash("Invalid session.", "error")
            return redirect(url_for("forgot_password"))

        entered_otp = request.form.get("otp", "").strip()
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")

        if new_password != confirm_password:
            flash("Passwords do not match.", "error")
            return redirect(url_for("reset_password"))

        if len(new_password) < 8:
            flash("Password must be at least 8 characters.", "error")
            return redirect(url_for("reset_password"))

        if not verify_otp(user, entered_otp, otp_type="reset"):
            flash("Invalid or expired OTP.", "error")
            return redirect(url_for("reset_password"))

        # OTP valid — update password and clear reset fields
        user.password_hash = hash_password(new_password)
        user.reset_otp_hash = None
        user.reset_otp_expiry = None
        user.failed_attempts = 0
        user.locked_until = None
        db.session.commit()

        session.pop("reset_user_id", None)

        log_action(f"Password reset completed for user id={user.id}")

        flash("Password reset successful. Please log in.", "success")

        return redirect(url_for("login"))

    return render_template("auth/reset_password.html")


# =========================================================
# CUSTOM ERROR HANDLERS
# =========================================================

@app.errorhandler(404)
def not_found(e):
    return render_template("errors/404.html"), 404


@app.errorhandler(500)
def server_error(e):
    db.session.rollback()
    return render_template("errors/500.html"), 500


@app.errorhandler(403)
def forbidden(e):
    return render_template("errors/403.html"), 403


# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():

    if session.get("user_id"):
        log_action("User logged out")

    session.clear()

    flash("Logged out successfully", "success")

    return redirect(url_for("login"))


# =========================================================
# DATABASE INIT
# =========================================================

with app.app_context():
    db.create_all()


# =========================================================
# RUN SERVER
# =========================================================

if __name__ == "__main__":
    HOST = os.environ.get("HOST", "0.0.0.0")
    PORT = int(os.environ.get("PORT", 6201))
    THREADS = int(os.environ.get("THREADS", 4))
    print(f"[waitress] Serving on http://{HOST}:{PORT} (threads={THREADS})")
    from waitress import serve
    serve(app, host=HOST, port=PORT, threads=THREADS)
