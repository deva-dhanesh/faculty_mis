from app import app, db
from sqlalchemy import text

with app.app_context():
    db.session.execute(text("""
        CREATE TABLE IF NOT EXISTS faculty_fellowships (
            id SERIAL PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id),
            award_name VARCHAR(500) NOT NULL,
            financial_support VARCHAR(100),
            grant_purpose VARCHAR(500),
            support_type VARCHAR(100),
            national_international VARCHAR(20),
            award_date DATE,
            awarding_agency VARCHAR(300),
            duration VARCHAR(100),
            research_topic VARCHAR(500),
            location VARCHAR(500),
            collaborating_institution VARCHAR(500),
            grant_letter_link VARCHAR(500),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """))
    db.session.commit()
    print("faculty_fellowships table created successfully.")
