#!/usr/bin/env python
"""
Test script to diagnose report generation issues
"""

import sys
import os

print("=" * 60)
print("REPORT GENERATION DIAGNOSTIC TEST")
print("=" * 60)

# Test 1: Check if all required modules are available
print("\n[1] Checking required modules...")
required_modules = ['plotly', 'reportlab', 'flask', 'sqlalchemy']

for module in required_modules:
    try:
        __import__(module)
        print(f"  [OK] {module}")
    except ImportError as e:
        print(f"  [FAIL] {module}: {e}")

# Test 2: Check report_generation module
print("\n[2] Checking report_generation module...")
try:
    from report_generation import (
        generate_charts, generate_interpretation,
        compile_summary, generate_detailed_stats, generate_pdf_report
    )
    print("  [OK] All functions imported successfully")
except Exception as e:
    print(f"  [FAIL] Import error: {e}")
    sys.exit(1)

# Test 3: Check models
print("\n[3] Checking database models...")
try:
    from models import User, FacultyProfile
    print("  [OK] User model")
    print("  [OK] FacultyProfile model")
except Exception as e:
    print(f"  [FAIL] Model error: {e}")
    sys.exit(1)

# Test 4: Check if User has expected attributes
print("\n[4] Checking User model attributes...")
try:
    user_attrs = ['id', 'email', 'employee_id', 'role']
    for attr in user_attrs:
        if hasattr(User, attr) or attr == 'id':
            print(f"  [OK] User.{attr}")
        else:
            print(f"  [WARN] User.{attr} not directly found")
except Exception as e:
    print(f"  [FAIL] {e}")

# Test 5: Check if FacultyProfile has full_name
print("\n[5] Checking FacultyProfile model attributes...")
try:
    fp_attrs = ['user_id', 'full_name', 'employee_id']
    for attr in fp_attrs:
        if hasattr(FacultyProfile, attr):
            print(f"  [OK] FacultyProfile.{attr}")
        else:
            print(f"  [WARN] FacultyProfile.{attr} not found")
except Exception as e:
    print(f"  [FAIL] {e}")

# Test 6: Check temp_reports directory
print("\n[6] Checking temp_reports directory...")
if not os.path.exists('temp_reports'):
    try:
        os.makedirs('temp_reports')
        print("  [OK] Created temp_reports directory")
    except Exception as e:
        print(f"  [FAIL] Cannot create directory: {e}")
else:
    print("  [OK] temp_reports directory exists")

# Test 7: Check templates
print("\n[7] Checking report templates...")
templates = [
    'templates/faculty/generate_report.html',
    'templates/faculty/view_report.html'
]
for template in templates:
    if os.path.exists(template):
        print(f"  [OK] {template}")
    else:
        print(f"  [FAIL] {template} not found")

print("\n" + "=" * 60)
print("DIAGNOSTIC TEST COMPLETE")
print("=" * 60)
print("\nIf all tests passed, the report generation should work!")
