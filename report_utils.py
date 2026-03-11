from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from models import FacultyProfile, FacultyPublication, FacultyProject
from datetime import datetime


W, H = A4
MARGIN = 20 * mm


def _header(c, title, subtitle=""):
    c.setFillColor(colors.HexColor("#1e293b"))
    c.rect(0, H - 22 * mm, W, 22 * mm, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(MARGIN, H - 14 * mm, title)
    if subtitle:
        c.setFont("Helvetica", 9)
        c.drawString(MARGIN, H - 19 * mm, subtitle)
    c.setFillColor(colors.HexColor("#64748b"))
    c.setFont("Helvetica", 8)
    c.drawString(MARGIN, 10 * mm, f"Generated on {datetime.now().strftime('%d %b %Y, %H:%M')}")
    c.drawRightString(W - MARGIN, 10 * mm, "Faculty MIS — Jain University")


def _section(c, y, label):
    c.setFillColor(colors.HexColor("#f1f5f9"))
    c.rect(MARGIN, y - 5, W - 2 * MARGIN, 16, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#0f172a"))
    c.setFont("Helvetica-Bold", 9)
    c.drawString(MARGIN + 4, y, label)
    return y - 22


def _row(c, y, label, value, alt=False):
    if alt:
        c.setFillColor(colors.HexColor("#f8fafc"))
        c.rect(MARGIN, y - 4, W - 2 * MARGIN, 14, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#64748b"))
    c.setFont("Helvetica", 8)
    c.drawString(MARGIN + 4, y, label)
    c.setFillColor(colors.HexColor("#0f172a"))
    c.setFont("Helvetica", 8)
    c.drawString(MARGIN + 55 * mm, y, str(value) if value else "—")
    return y - 16


def _check_page(c, y, title, subtitle=""):
    if y < 35 * mm:
        c.showPage()
        _header(c, title, subtitle)
        return H - 35 * mm
    return y


# =========================================================
# ADMIN REPORT — all faculty summary
# =========================================================

def generate_admin_report(filepath):
    c = canvas.Canvas(filepath, pagesize=A4)
    title = "Faculty MIS — All Faculty Report"
    subtitle = f"Generated: {datetime.now().strftime('%d %B %Y')}"
    _header(c, title, subtitle)
    y = H - 35 * mm

    profiles = FacultyProfile.query.order_by(
        FacultyProfile.department, FacultyProfile.full_name
    ).all()

    # Table header
    c.setFillColor(colors.HexColor("#1e293b"))
    c.rect(MARGIN, y - 5, W - 2 * MARGIN, 16, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 8)
    pcols = [MARGIN+4, MARGIN+8*mm, MARGIN+55*mm, MARGIN+97*mm, MARGIN+127*mm, MARGIN+146*mm, MARGIN+159*mm]
    for txt, x in zip(["#","Name","Department","Designation","Exp","Pubs","Projects"], pcols):
        c.drawString(x, y, txt)
    y -= 20

    for idx, p in enumerate(profiles):
        y = _check_page(c, y, title, subtitle)
        if idx % 2 == 0:
            c.setFillColor(colors.HexColor("#f8fafc"))
            c.rect(MARGIN, y - 4, W - 2 * MARGIN, 14, fill=1, stroke=0)
        pub_count  = FacultyPublication.query.filter_by(user_id=p.user_id).count()
        proj_count = FacultyProject.query.filter_by(user_id=p.user_id).count()
        c.setFillColor(colors.HexColor("#64748b"))
        c.setFont("Helvetica", 8)
        c.drawString(pcols[0], y, str(idx + 1))
        c.setFillColor(colors.HexColor("#0f172a"))
        c.setFont("Helvetica", 8)
        c.drawString(pcols[1], y, (p.full_name or "")[:28])
        c.drawString(pcols[2], y, (p.department or "")[:22])
        c.drawString(pcols[3], y, (p.designation or "")[:18])
        c.drawString(pcols[4], y, str(p.experience_years or "—"))
        c.drawString(pcols[5], y, str(pub_count))
        c.drawString(pcols[6], y, str(proj_count))
        y -= 16

    c.save()


# Keep old name as alias so existing import doesn't break
generate_faculty_report = generate_admin_report


# =========================================================
# FACULTY PERSONAL REPORT — one faculty member
# =========================================================

def generate_personal_report(filepath, user_id):
    c = canvas.Canvas(filepath, pagesize=A4)

    profile      = FacultyProfile.query.filter_by(user_id=user_id).first()
    publications = FacultyPublication.query.filter_by(user_id=user_id).order_by(
        FacultyPublication.publication_date.desc()).all()
    projects     = FacultyProject.query.filter_by(user_id=user_id).order_by(
        FacultyProject.date_of_award.desc()).all()

    name     = profile.full_name if profile else "Faculty Member"
    dept     = profile.department if profile else ""
    title    = f"Academic Report — {name}"
    subtitle = dept or ""

    _header(c, title, subtitle)
    y = H - 35 * mm

    # ── Profile ──────────────────────────────────────────
    y = _section(c, y, "PROFILE INFORMATION")
    if profile:
        fields = [
            ("Employee ID",      profile.employee_id),
            ("Full Name",        profile.full_name),
            ("Designation",      profile.designation),
            ("Department",       profile.department),
            ("Qualification",    profile.qualification),
            ("Specialization",   profile.specialization),
            ("Experience",       f"{profile.experience_years} years" if profile.experience_years else None),
            ("Mobile",           profile.mobile),
            ("University Email", profile.email_university),
            ("Personal Email",   profile.email_personal),
            ("Date of Joining",  profile.date_of_joining.strftime("%d %b %Y") if profile.date_of_joining else None),
        ]
        for i, (lbl, val) in enumerate(fields):
            y = _check_page(c, y, title, subtitle)
            y = _row(c, y, lbl, val, alt=(i % 2 == 0))
    else:
        c.setFillColor(colors.HexColor("#64748b"))
        c.setFont("Helvetica-Oblique", 9)
        c.drawString(MARGIN + 4, y, "No profile information available.")
        y -= 20
    y -= 6

    # ── Publications ─────────────────────────────────────
    y = _check_page(c, y, title, subtitle)
    y = _section(c, y, f"PUBLICATIONS  ({len(publications)} total)")
    if publications:
        c.setFillColor(colors.HexColor("#e2e8f0"))
        c.rect(MARGIN, y - 5, W - 2 * MARGIN, 14, fill=1, stroke=0)
        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 7.5)
        xcols = [MARGIN+4, MARGIN+6*mm, MARGIN+70*mm, MARGIN+100*mm, MARGIN+120*mm, MARGIN+140*mm]
        for txt, x in zip(["#","Title","Journal","Date","Type","Indexing"], xcols):
            c.drawString(x, y, txt)
        y -= 18
        for idx, pub in enumerate(publications):
            y = _check_page(c, y, title, subtitle)
            if idx % 2 == 0:
                c.setFillColor(colors.HexColor("#f8fafc"))
                c.rect(MARGIN, y - 4, W - 2 * MARGIN, 14, fill=1, stroke=0)
            c.setFillColor(colors.HexColor("#64748b"))
            c.setFont("Helvetica", 7.5)
            c.drawString(xcols[0], y, str(idx + 1))
            c.setFillColor(colors.HexColor("#0f172a"))
            c.setFont("Helvetica", 7.5)
            c.drawString(xcols[1], y, (pub.title or "")[:45])
            c.drawString(xcols[2], y, (pub.journal or "—")[:20])
            c.drawString(xcols[3], y, pub.publication_date.strftime("%d-%b-%Y") if pub.publication_date else "—")
            c.drawString(xcols[4], y, pub.publication_type or "—")
            c.drawString(xcols[5], y, pub.indexing or "—")
            y -= 15
    else:
        c.setFillColor(colors.HexColor("#64748b"))
        c.setFont("Helvetica-Oblique", 9)
        c.drawString(MARGIN + 4, y, "No publications on record.")
        y -= 20
    y -= 6

    # ── Projects ─────────────────────────────────────────
    y = _check_page(c, y, title, subtitle)
    y = _section(c, y, f"RESEARCH PROJECTS  ({len(projects)} total)")
    if projects:
        c.setFillColor(colors.HexColor("#e2e8f0"))
        c.rect(MARGIN, y - 5, W - 2 * MARGIN, 14, fill=1, stroke=0)
        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 7.5)
        rcols = [MARGIN+4, MARGIN+6*mm, MARGIN+75*mm, MARGIN+112*mm, MARGIN+132*mm, MARGIN+153*mm]
        for txt, x in zip(["#","Scheme / Project Name","Funding Agency","Type","Lakhs","Status"], rcols):
            c.drawString(x, y, txt)
        y -= 18
        for idx, proj in enumerate(projects):
            y = _check_page(c, y, title, subtitle)
            if idx % 2 == 0:
                c.setFillColor(colors.HexColor("#f8fafc"))
                c.rect(MARGIN, y - 4, W - 2 * MARGIN, 14, fill=1, stroke=0)
            c.setFillColor(colors.HexColor("#64748b"))
            c.setFont("Helvetica", 7.5)
            c.drawString(rcols[0], y, str(idx + 1))
            c.setFillColor(colors.HexColor("#0f172a"))
            c.setFont("Helvetica", 7.5)
            c.drawString(rcols[1], y, (proj.scheme_name or "")[:48])
            c.drawString(rcols[2], y, (proj.funding_agency or "—")[:20])
            c.drawString(rcols[3], y, proj.project_type or "—")
            c.drawString(rcols[4], y, f"{proj.amount:.2f}" if proj.amount else "—")
            c.drawString(rcols[5], y, proj.status or "—")
            y -= 15
    else:
        c.setFillColor(colors.HexColor("#64748b"))
        c.setFont("Helvetica-Oblique", 9)
        c.drawString(MARGIN + 4, y, "No projects on record.")
        y -= 20

    c.save()
