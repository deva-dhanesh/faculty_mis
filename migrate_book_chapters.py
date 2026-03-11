"""
migrate_book_chapters.py
Creates the faculty_book_chapters table in the database.
Run once, then delete this file.
"""
from app import app, db
from models import FacultyBookChapter
from sqlalchemy import text

with app.app_context():
    # Create the table if it doesn't exist
    db.session.execute(text("""
        CREATE TABLE IF NOT EXISTS faculty_book_chapters (
            id                     SERIAL PRIMARY KEY,
            user_id                INTEGER NOT NULL REFERENCES users(id),
            book_title             VARCHAR(500),
            chapter_title          VARCHAR(500),
            book_or_chapter        VARCHAR(20),
            translated_title       VARCHAR(500),
            proceedings_publisher  VARCHAR(500),
            internal_external      VARCHAR(20),
            national_international VARCHAR(20),
            publication_date       DATE,
            isbn                   VARCHAR(100),
            co_authors             VARCHAR(500),
            doi                    VARCHAR(200),
            indexed_in             VARCHAR(300),
            journal_link           VARCHAR(500),
            supporting_doc_link    VARCHAR(500),
            created_at             TIMESTAMP DEFAULT NOW()
        );
    """))
    db.session.commit()
    print("✅  faculty_book_chapters table created successfully!")
