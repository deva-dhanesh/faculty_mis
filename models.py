from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()


# =========================================================
# USER TABLE (Authentication & Login)
# =========================================================
class User(UserMixin, db.Model):

    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)

    employee_id = db.Column(
        db.String(50),
        unique=True,
        nullable=False
    )

    email = db.Column(
        db.String(120),
        unique=True,
        nullable=False
    )

    phone = db.Column(
        db.String(15),
        nullable=False
    )

    password_hash = db.Column(
        db.String(255),
        nullable=False
    )

    role = db.Column(
        db.String(20),
        nullable=False
    )

    otp_hash = db.Column(db.String(255))
    otp_expiry = db.Column(db.DateTime)

    reset_otp_hash = db.Column(db.String(255))
    reset_otp_expiry = db.Column(db.DateTime)

    # Account lockout
    failed_attempts = db.Column(db.Integer, default=0, nullable=False)
    locked_until = db.Column(db.DateTime, nullable=True)

    is_active_account = db.Column(db.Boolean, default=True, nullable=False)

    # OTP brute-force
    otp_attempts = db.Column(db.Integer, default=0, nullable=False)

    # First-login flag — faculty must reset password on first login
    must_reset_password = db.Column(db.Boolean, default=False, nullable=True)

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )

    def get_id(self):
        return str(self.id)


# =========================================================
# FACULTY PROFILE TABLE (Core Faculty Information)
# =========================================================
class FacultyProfile(db.Model):

    __tablename__ = "faculty_profiles"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False,
        unique=True
    )

    employee_id = db.Column(db.String(50))

    full_name = db.Column(db.String(200))

    pan = db.Column(db.String(20))

    designation = db.Column(db.String(200))

    date_of_joining = db.Column(db.Date)

    date_of_birth = db.Column(db.Date)

    appointment_nature = db.Column(db.String(100))

    qualification = db.Column(db.String(200))

    department = db.Column(db.String(200))

    experience_years = db.Column(db.Integer)

    appointment_letter_url = db.Column(db.String(500))

    mobile = db.Column(db.String(20))

    email_personal = db.Column(db.String(200))

    email_university = db.Column(db.String(200))

    specialization = db.Column(db.String(300))

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )

    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow
    )


# =========================================================
# FACULTY RESIGNATION TABLE
# =========================================================
class FacultyResignation(db.Model):

    __tablename__ = "faculty_resignations"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    resignation_date = db.Column(db.Date)

    reason = db.Column(db.Text)

    relieving_date = db.Column(db.Date)

    resignation_letter_url = db.Column(db.String(500))


# =========================================================
# FACULTY DEGREE TABLE (Multiple Degrees Allowed)
# =========================================================
class FacultyDegree(db.Model):

    __tablename__ = "faculty_degrees"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    degree_name = db.Column(db.String(200))

    degree_start_date = db.Column(db.Date)

    degree_award_date = db.Column(db.Date)

    university = db.Column(db.String(300))

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )


# =========================================================
# FACULTY DOCUMENT TABLE (NEW - Secure Document Storage)
# =========================================================
class FacultyDocument(db.Model):

    __tablename__ = "faculty_documents"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    document_type = db.Column(
        db.String(100),
        nullable=False
    )

    file_name = db.Column(
        db.String(300),
        nullable=False
    )

    file_path = db.Column(
        db.String(500),
        nullable=False
    )

    uploaded_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )

# =========================================================
# FACULTY PUBLICATIONS TABLE
# =========================================================
class FacultyPublication(db.Model):

    __tablename__ = "faculty_publications"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    # Position of the faculty as Author (First / Second / Corresponding / etc.)
    author_position = db.Column(db.String(200))

    # Scopus ID / ORCID ID / Google Scholar ID
    scholar_id = db.Column(db.String(300))

    # Title of the Paper / Conference Proceeding
    title = db.Column(
        db.String(500),
        nullable=False
    )

    # Name of Journal
    journal = db.Column(db.String(500))

    # Date of Publication (stored as Date)
    publication_date = db.Column(db.Date)

    # ISSN / ISBN Number
    issn_isbn = db.Column(db.String(100))

    # h-index
    h_index = db.Column(db.String(50))

    # Citation index
    citation_index = db.Column(db.String(100))

    # Journal Quartile (Q1 / Q2 / Q3 / Q4)
    journal_quartile = db.Column(db.String(10))

    # Type of Publication (Journal / Conference)
    publication_type = db.Column(db.String(50))

    # Impact Factor
    impact_factor = db.Column(db.String(50))

    # Indexing (Scopus / WoS / UGC CARE / PubMed / Others)
    indexing = db.Column(db.String(200))

    # DOI
    doi = db.Column(db.String(300))

    # Link to Article / Conference Proceeding
    article_link = db.Column(db.String(500))

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )

# =========================================================
# FACULTY RESEARCH PROJECTS TABLE
# =========================================================
class FacultyProject(db.Model):

    __tablename__ = "faculty_projects"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    # Name of the Scheme / Project / Endowments / Chairs
    scheme_name = db.Column(db.String(500), nullable=False)

    # Name of the Principal Investigator / Co-Investigator
    pi_co_pi = db.Column(db.String(500))

    funding_agency = db.Column(db.String(300))

    project_type = db.Column(db.String(100))   # Govt. / Non-Govt.

    department = db.Column(db.String(300))

    date_of_award = db.Column(db.Date)

    amount = db.Column(db.Float)               # INR in Lakhs

    duration_years = db.Column(db.String(50))  # e.g. "2", "2.5"

    status = db.Column(db.String(100))         # Completed / Ongoing

    date_of_completion = db.Column(db.Date)

    objectives = db.Column(db.Text)

    collaborating_institutions = db.Column(db.String(500))

    document_link = db.Column(db.String(500))  # URL to Sanction Order

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

# =========================================================
# VALUE ADDED PROGRAMS — COURSES ATTENDED
# =========================================================
class CourseAttended(db.Model):

    __tablename__ = "courses_attended"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    course_name = db.Column(db.String(500), nullable=False)

    online_course_name = db.Column(db.String(500))

    date_from = db.Column(db.Date)

    date_to = db.Column(db.Date)

    mode = db.Column(db.String(300))          # Offline / Online + platform

    contact_hours = db.Column(db.Float)

    offered_by = db.Column(db.String(300))    # University/Institute/Industry

    certificate_link = db.Column(db.String(500))

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )


# =========================================================
# VALUE ADDED PROGRAMS — COURSES OFFERED
# =========================================================
class CourseOffered(db.Model):

    __tablename__ = "courses_offered"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    course_name = db.Column(db.String(500), nullable=False)

    online_course_name = db.Column(db.String(500))

    credits_assigned = db.Column(db.String(300))  # Yes/No + credits value

    program_name = db.Column(db.String(300))

    department = db.Column(db.String(200))

    date_from = db.Column(db.Date)

    date_to = db.Column(db.Date)

    times_offered = db.Column(db.Integer)

    mode = db.Column(db.String(300))

    contact_hours = db.Column(db.Float)

    students_enrolled_link = db.Column(db.String(500))

    students_completing_link = db.Column(db.String(500))

    attendance_link = db.Column(db.String(500))

    brochure_link = db.Column(db.String(500))

    certificate_link = db.Column(db.String(500))

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )


# =========================================================
# AWARDS
# =========================================================
class FacultyAward(db.Model):

    __tablename__ = "faculty_awards"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    title = db.Column(db.String(500), nullable=False)          # Title of Innovation / Academic Achievement

    nature_of_award = db.Column(db.String(100))                # Academic/Research/Innovations/Sports/Cultural/Alumni/Women Cell

    event_level = db.Column(db.String(100))                    # Local/State/National/International

    date_of_award = db.Column(db.Date)

    category = db.Column(db.String(100))                       # Department/Teachers/Research Scholars

    awarding_agency = db.Column(db.String(300))

    award_amount = db.Column(db.String(200))                   # optional text e.g. "₹5000" or "—"

    research_area = db.Column(db.String(300))

    collaborators = db.Column(db.String(500))

    document_link = db.Column(db.String(500))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# =========================================================
# BOOK CHAPTER TABLE
# =========================================================
class FacultyBookChapter(db.Model):

    __tablename__ = "faculty_book_chapters"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    book_title = db.Column(db.String(500))

    chapter_title = db.Column(db.String(500))

    book_or_chapter = db.Column(db.String(20))          # Book / Chapter

    translated_title = db.Column(db.String(500))

    proceedings_publisher = db.Column(db.String(500))

    internal_external = db.Column(db.String(20))        # Internal / External

    national_international = db.Column(db.String(20))   # National / International

    publication_date = db.Column(db.Date, nullable=True)

    isbn = db.Column(db.String(100))

    co_authors = db.Column(db.String(500))

    doi = db.Column(db.String(200))

    indexed_in = db.Column(db.String(300))

    journal_link = db.Column(db.String(500))

    supporting_doc_link = db.Column(db.String(500))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# =========================================================
# GUEST LECTURE TABLE
# =========================================================
class FacultyGuestLecture(db.Model):

    __tablename__ = "faculty_guest_lectures"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    lecture_title = db.Column(db.String(500), nullable=False)

    organization_location = db.Column(db.String(500))

    lecture_date = db.Column(db.Date, nullable=True)

    jain_or_outside = db.Column(db.String(20))       # JAIN / Outside

    mode = db.Column(db.String(20))                   # In-person / Online

    audience_type = db.Column(db.String(30))           # Students / Faculty / Industry

    brochure_link = db.Column(db.String(500))

    supporting_doc_link = db.Column(db.String(500))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# =========================================================
# PATENT / IP TABLE
# =========================================================
class FacultyPatent(db.Model):

    __tablename__ = "faculty_patents"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    application_number = db.Column(db.String(100))

    ip_type = db.Column(db.String(40))  # Patent / Copyright / Trademark / GI / Design Registration

    title = db.Column(db.String(500), nullable=False)

    status = db.Column(db.String(20))  # Filed / Published / Awarded

    filing_date = db.Column(db.Date, nullable=True)

    published_date = db.Column(db.Date, nullable=True)

    grant_date = db.Column(db.Date, nullable=True)

    awarding_agency = db.Column(db.String(300))

    national_international = db.Column(db.String(20))  # National / International

    commercialization_details = db.Column(db.Text)

    oer_contribution = db.Column(db.Text)

    supporting_doc_link = db.Column(db.String(500))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# =========================================================
# FELLOWSHIP TABLE
# =========================================================
class FacultyFellowship(db.Model):

    __tablename__ = "faculty_fellowships"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    award_name = db.Column(db.String(500), nullable=False)

    financial_support = db.Column(db.String(100))        # INR amount

    grant_purpose = db.Column(db.String(500))

    support_type = db.Column(db.String(100))

    national_international = db.Column(db.String(20))    # National / International

    award_date = db.Column(db.Date, nullable=True)

    awarding_agency = db.Column(db.String(300))

    duration = db.Column(db.String(100))

    research_topic = db.Column(db.String(500))

    location = db.Column(db.String(500))

    collaborating_institution = db.Column(db.String(500))

    grant_letter_link = db.Column(db.String(500))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# =========================================================
# CONFERENCES — PARTICIPATED
# =========================================================
class ConferenceParticipated(db.Model):

    __tablename__ = "conferences_participated"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    conference_title = db.Column(db.String(500), nullable=False)

    paper_title = db.Column(db.String(500))

    organisers = db.Column(db.String(500))

    leads_to_publication = db.Column(db.String(10))

    paper_link = db.Column(db.String(500))

    collaboration = db.Column(db.String(500))

    focus_area = db.Column(db.String(200))

    date_from = db.Column(db.Date)

    date_to = db.Column(db.Date)

    national_international = db.Column(db.String(20))

    proceedings_indexed = db.Column(db.String(10))

    indexing_details = db.Column(db.String(500))

    certificate_link = db.Column(db.String(500))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# =========================================================
# CONFERENCES — ORGANISED
# =========================================================
class ConferenceOrganised(db.Model):

    __tablename__ = "conferences_organised"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    title = db.Column(db.String(500), nullable=False)

    department = db.Column(db.String(300))

    faculty_role = db.Column(db.String(300))

    collaboration = db.Column(db.String(500))

    focus_area = db.Column(db.String(200))

    national_international = db.Column(db.String(20))

    num_participants = db.Column(db.Integer)

    date_from = db.Column(db.Date)

    date_to = db.Column(db.Date)

    proceedings_indexed = db.Column(db.String(10))

    indexing_details = db.Column(db.String(500))

    activity_report_link = db.Column(db.String(500))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# =========================================================
# FDP — PARTICIPATED
# =========================================================
class FDPParticipated(db.Model):

    __tablename__ = "fdp_participated"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    program_title = db.Column(db.String(500), nullable=False)

    duration_days = db.Column(db.String(50))

    start_date = db.Column(db.Date)

    end_date = db.Column(db.Date)

    program_type = db.Column(db.String(200))

    national_international = db.Column(db.String(20))

    organizing_agency = db.Column(db.String(500))

    location = db.Column(db.String(500))

    mode = db.Column(db.String(50))

    funding = db.Column(db.String(100))

    certificate_link = db.Column(db.String(500))

    brochure_link = db.Column(db.String(500))

    enrolled_coursera = db.Column(db.String(10))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# =========================================================
# FDP — ORGANISED
# =========================================================
class FDPOrganised(db.Model):

    __tablename__ = "fdp_organised"

    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    title = db.Column(db.String(500), nullable=False)

    department = db.Column(db.String(300))

    faculty_role = db.Column(db.String(300))

    collaboration = db.Column(db.String(500))

    focus_area = db.Column(db.String(200))

    national_international = db.Column(db.String(20))

    num_participants = db.Column(db.Integer)

    date_from = db.Column(db.Date)

    date_to = db.Column(db.Date)

    activity_report_link = db.Column(db.String(500))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# =========================================================
# AUDIT LOG TABLE (Tracks all system actions)
# =========================================================
class AuditLog(db.Model):

    __tablename__ = "audit_logs"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=True
    )

    user_email = db.Column(
        db.String(200)
    )

    role = db.Column(
        db.String(50)
    )

    action = db.Column(
        db.String(300),
        nullable=False
    )

    ip_address = db.Column(
        db.String(100)
    )

    timestamp = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )
