"""
One-time migration: update faculty_publications table to new schema.
Run once:  python migrate_publications.py
"""
from app import app, db
from sqlalchemy import text

ADDS = [
    "ALTER TABLE faculty_publications ADD COLUMN IF NOT EXISTS author_position VARCHAR(200)",
    "ALTER TABLE faculty_publications ADD COLUMN IF NOT EXISTS scholar_id VARCHAR(300)",
    "ALTER TABLE faculty_publications ADD COLUMN IF NOT EXISTS publication_date DATE",
    "ALTER TABLE faculty_publications ADD COLUMN IF NOT EXISTS issn_isbn VARCHAR(100)",
    "ALTER TABLE faculty_publications ADD COLUMN IF NOT EXISTS h_index VARCHAR(50)",
    "ALTER TABLE faculty_publications ADD COLUMN IF NOT EXISTS citation_index VARCHAR(100)",
    "ALTER TABLE faculty_publications ADD COLUMN IF NOT EXISTS journal_quartile VARCHAR(10)",
    "ALTER TABLE faculty_publications ADD COLUMN IF NOT EXISTS publication_type VARCHAR(50)",
    "ALTER TABLE faculty_publications ADD COLUMN IF NOT EXISTS impact_factor VARCHAR(50)",
    "ALTER TABLE faculty_publications ADD COLUMN IF NOT EXISTS article_link VARCHAR(500)",
]

# Migrate old publication_year → publication_date (set Jan 1 of that year)
MIGRATE_YEAR = """
    UPDATE faculty_publications
    SET publication_date = make_date(publication_year, 1, 1)
    WHERE publication_year IS NOT NULL AND publication_date IS NULL
"""

DROPS = [
    "ALTER TABLE faculty_publications DROP COLUMN IF EXISTS authors",
    "ALTER TABLE faculty_publications DROP COLUMN IF EXISTS publication_year",
    "ALTER TABLE faculty_publications DROP COLUMN IF EXISTS document_path",
]

with app.app_context():
    conn = db.engine.connect()
    trans = conn.begin()
    try:
        for sql in ADDS:
            print(f"  + {sql.split('ADD COLUMN IF NOT EXISTS ')[-1].split()[0]}")
            conn.execute(text(sql))

        print("  ~ Migrating publication_year → publication_date ...")
        conn.execute(text(MIGRATE_YEAR))

        for sql in DROPS:
            col = sql.split("DROP COLUMN IF EXISTS ")[-1]
            print(f"  - {col}")
            conn.execute(text(sql))

        trans.commit()
        print("\n✔ Migration completed successfully.")
    except Exception as e:
        trans.rollback()
        print(f"\n✖ Migration failed: {e}")
    finally:
        conn.close()
