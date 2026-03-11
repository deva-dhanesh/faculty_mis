from app import app, db
from sqlalchemy import text

with app.app_context():
    db.session.execute(text("""
        CREATE TABLE IF NOT EXISTS conferences_participated (
            id SERIAL PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id),
            conference_title VARCHAR(500) NOT NULL,
            paper_title VARCHAR(500),
            organisers VARCHAR(500),
            leads_to_publication VARCHAR(10),
            paper_link VARCHAR(500),
            collaboration VARCHAR(500),
            focus_area VARCHAR(200),
            date_from DATE,
            date_to DATE,
            national_international VARCHAR(20),
            proceedings_indexed VARCHAR(10),
            indexing_details VARCHAR(500),
            certificate_link VARCHAR(500),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """))
    db.session.execute(text("""
        CREATE TABLE IF NOT EXISTS conferences_organised (
            id SERIAL PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id),
            title VARCHAR(500) NOT NULL,
            department VARCHAR(300),
            faculty_role VARCHAR(300),
            collaboration VARCHAR(500),
            focus_area VARCHAR(200),
            national_international VARCHAR(20),
            num_participants INTEGER,
            date_from DATE,
            date_to DATE,
            proceedings_indexed VARCHAR(10),
            indexing_details VARCHAR(500),
            activity_report_link VARCHAR(500),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """))
    db.session.commit()
    print("conferences_participated table created successfully.")
    print("conferences_organised table created successfully.")
