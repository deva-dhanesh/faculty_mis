"""Create FDP tables: fdp_participated and fdp_organised."""

from app import app, db
from sqlalchemy import text

with app.app_context():
    db.session.execute(text("""
        CREATE TABLE IF NOT EXISTS fdp_participated (
            id              SERIAL PRIMARY KEY,
            user_id         INTEGER NOT NULL REFERENCES users(id),
            program_title   VARCHAR(500) NOT NULL,
            duration_days   VARCHAR(50),
            start_date      DATE,
            end_date        DATE,
            program_type    VARCHAR(200),
            national_international VARCHAR(20),
            organizing_agency VARCHAR(500),
            location        VARCHAR(500),
            mode            VARCHAR(50),
            funding         VARCHAR(100),
            certificate_link VARCHAR(500),
            brochure_link   VARCHAR(500),
            enrolled_coursera VARCHAR(10),
            created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """))

    db.session.execute(text("""
        CREATE TABLE IF NOT EXISTS fdp_organised (
            id              SERIAL PRIMARY KEY,
            user_id         INTEGER NOT NULL REFERENCES users(id),
            title           VARCHAR(500) NOT NULL,
            department      VARCHAR(300),
            faculty_role    VARCHAR(300),
            collaboration   VARCHAR(500),
            focus_area      VARCHAR(200),
            national_international VARCHAR(20),
            num_participants INTEGER,
            date_from       DATE,
            date_to         DATE,
            activity_report_link VARCHAR(500),
            created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """))

    db.session.commit()
    print("✅ FDP tables created successfully.")
