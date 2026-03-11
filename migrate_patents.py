"""Create faculty_patents table."""
from app import app, db
from sqlalchemy import text

with app.app_context():
    db.session.execute(text("""
        CREATE TABLE IF NOT EXISTS faculty_patents (
            id                      SERIAL PRIMARY KEY,
            user_id                 INTEGER NOT NULL REFERENCES users(id),
            application_number      VARCHAR(100),
            ip_type                 VARCHAR(40),
            title                   VARCHAR(500) NOT NULL,
            status                  VARCHAR(20),
            filing_date             DATE,
            published_date          DATE,
            grant_date              DATE,
            awarding_agency         VARCHAR(300),
            national_international  VARCHAR(20),
            commercialization_details TEXT,
            oer_contribution        TEXT,
            supporting_doc_link     VARCHAR(500),
            created_at              TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """))
    db.session.commit()
    print("✅  faculty_patents table ready.")
