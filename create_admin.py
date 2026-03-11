from app import app
from models import db, User
from security_utils import hash_password
import sys

def get_input(prompt, default):
    """Get input from command line args, stdin, or use default"""
    user_input = input(f"{prompt} [{default}]: ").strip()
    return user_input if user_input else default

def show_help():
    """Display usage information"""
    print("Usage: python create_admin.py [email] [phone] [password]")
    print("\nParameters:")
    print("  email    Admin email address (default: admin@mis.edu)")
    print("  phone    Admin phone number (default: 9999999999)")
    print("  password Admin password (default: Admin@123)")
    print("\nExample:")
    print("  python create_admin.py admin@example.com 1234567890 MyPassword@123")
    sys.exit(0)

if len(sys.argv) > 1 and sys.argv[1] == "--help":
    show_help()

with app.app_context():
    # Try to get from command line args, otherwise prompt user
    email = sys.argv[1] if len(sys.argv) > 1 else get_input("Email", "admin@mis.edu")
    phone = sys.argv[2] if len(sys.argv) > 2 else get_input("Phone", "9999999999")
    password = sys.argv[3] if len(sys.argv) > 3 else get_input("Password", "Admin@123")

    admin = User(
        employee_id="ADMIN001",
        email=email,
        phone=phone,
        password_hash=hash_password(password),
        role="admin"
    )

    db.session.add(admin)
    db.session.commit()

    print("Admin created")
