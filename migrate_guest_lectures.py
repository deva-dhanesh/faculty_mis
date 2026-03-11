"""Create the faculty_guest_lectures table if it does not exist."""

from app import app, db
from sqlalchemy import text

SQL = """
CREATE TABLE IF NOT EXISTS faculty_guest_lectures (
    id              SERIAL PRIMARY KEY,
    user_id         INTEGER NOT NULL REFERENCES users(id),
    lecture_title   VARCHAR(500) NOT NULL,
    organization_location VARCHAR(500),
    lecture_date    DATE,
    jain_or_outside VARCHAR(20),
    mode            VARCHAR(20),
    audience_type   VARCHAR(30),
    brochure_link   VARCHAR(500),
    supporting_doc_link VARCHAR(500),
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
"""

with app.app_context():
    db.session.execute(text(SQL))
    db.session.commit()
    print("✅  faculty_guest_lectures table ready.")
