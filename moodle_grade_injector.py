#!/usr/bin/env python3
"""
Moodle Grade Injector version4 - Complete version with google chrome profile selection
Better use custom profile

Usage: python moodle_grade_injector.py <grades.csv/xlsx> "moodle_quick_grading_url"

       python moodle_grade_injector.py homework1_grades_sample.csv "https://moodle.innopolis.university/mod/assign/view.php?id=131686&action=grading"
"""

import sys
import subprocess
import importlib
import os
import platform

# Auto-install required packages
def install_if_needed(package_name, import_name=None):
    """Install a package if it's not available"""
    if import_name is None:
        import_name = package_name

    try:
        importlib.import_module(import_name)
        return True
    except ImportError:
        print(f"📦 Installing {package_name}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name, "--quiet"])
            print(f"✓ {package_name} installed")
            return True
        except:
            print(f"❌ Failed to install {package_name}")
            return False

# Check and install required packages
print("🔍 Checking required packages...")
packages = [
    ("pandas", "pandas"),
    ("selenium", "selenium"),
    ("webdriver-manager", "webdriver_manager"),
    ("colorama", "colorama"),
    ("openpyxl", "openpyxl")
]

for package, import_name in packages:
    install_if_needed(package, import_name)

# Now import everything
import pandas as pd
import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from colorama import init, Fore, Style

# Initialize colorama for colored output
init()

def get_chrome_profiles():
    """Get list of available Chrome profiles including custom ones"""
    system = platform.system()
    profiles = []
    custom_profiles = []

    if system == "Linux":
        profile_path = os.path.expanduser("~/.config/google-chrome")
    elif system == "Windows":
        profile_path = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")
    elif system == "Darwin":  # macOS
        profile_path = os.path.expanduser("~/Library/Application Support/Google/Chrome")
    else:
        return profiles, custom_profiles

    if os.path.exists(profile_path):
        # Look for Profile directories
        for item in os.listdir(profile_path):
            if item == "Default" or item.startswith("Profile "):
                profiles.append(item)

    # Check for known custom profiles in home directory
    home = os.path.expanduser("~")
    custom_profile_names = [
        "moodle-grading",
        ".moodle-grading",
        "moodle_grading_profile",
        ".moodle_grading_profile",
        ".moodle_temp_profile"
    ]

    for profile_name in custom_profile_names:
        profile_path = os.path.join(home, profile_name)
        if os.path.exists(profile_path):
            custom_profiles.append(profile_path)

    # Also check for a saved profiles list
    config_file = os.path.join(home, ".moodle_grader_profiles.json")
    if os.path.exists(config_file):
        try:
            import json
            with open(config_file, 'r') as f:
                saved_profiles = json.load(f)
                for p in saved_profiles.get('custom_profiles', []):
                    if os.path.exists(p) and p not in custom_profiles:
                        custom_profiles.append(p)
        except:
            pass

    return profiles, custom_profiles

def save_custom_profile(profile_path):
    """Save custom profile path for future use"""
    config_file = os.path.expanduser("~/.moodle_grader_profiles.json")
    saved_profiles = {'custom_profiles': []}

    if os.path.exists(config_file):
        try:
            with open(config_file, 'r') as f:
                saved_profiles = json.load(f)
        except:
            pass

    if profile_path not in saved_profiles.get('custom_profiles', []):
        saved_profiles.setdefault('custom_profiles', []).append(profile_path)
        try:
            with open(config_file, 'w') as f:
                json.dump(saved_profiles, f, indent=2)
        except:
            pass

def select_chrome_profile():
    """Let user select which Chrome profile to use"""
    profiles, custom_profiles = get_chrome_profiles()

    print("\n" + "="*60)
    print("🌐 CHROME PROFILE SELECTION")
    print("="*60)

    option_num = 0
    options_map = {}

    print(f"\n{option_num}. 🆕 Create temporary profile (will need to log in)")
    options_map[option_num] = ('temp', None)
    option_num += 1

    # Show Chrome's built-in profiles
    if profiles:
        print("\n📂 Chrome Profiles:")
        for profile in profiles:
            if profile == "Default":
                print(f"{option_num}. 👤 Main Profile (Default)")
            else:
                print(f"{option_num}. 👤 {profile}")
            options_map[option_num] = ('profile', profile)
            option_num += 1

    # Show custom profiles
    if custom_profiles:
        print("\n📁 Custom Profiles (previously used):")
        for custom_path in custom_profiles:
            profile_name = os.path.basename(custom_path) or custom_path
            if "moodle" in profile_name.lower():
                icon = "🎓"
            else:
                icon = "📁"
            print(f"{option_num}. {icon} {profile_name}")
            if custom_path != profile_name:
                print(f"     Path: {custom_path}")
            options_map[option_num] = ('custom', custom_path)
            option_num += 1

    print(f"\n{option_num}. ➕ Enter new custom profile path")
    options_map[option_num] = ('new_custom', None)

    while True:
        try:
            choice = input(f"\nSelect profile option (0-{option_num}): ").strip()

            # Also allow typing profile names directly
            if not choice.isdigit():
                # Check if it matches a known profile name
                if choice == "moodle-grading" or choice in [os.path.basename(p) for p in custom_profiles]:
                    # Find the full path
                    for path in custom_profiles:
                        if os.path.basename(path) == choice or path == choice:
                            print(f"→ Using custom profile: {choice}")
                            return None, path
                    # If not found in custom_profiles, create the path
                    custom_path = os.path.expanduser(f"~/{choice}")
                    print(f"→ Using custom profile: {custom_path}")
                    save_custom_profile(custom_path)
                    return None, custom_path

            choice_num = int(choice)

            if choice_num in options_map:
                option_type, option_value = options_map[choice_num]

                if option_type == 'temp':
                    print("→ Using temporary profile")
                    return None, None
                elif option_type == 'profile':
                    print(f"→ Selected: {option_value}")
                    return option_value, None
                elif option_type == 'custom':
                    print(f"→ Using custom profile: {os.path.basename(option_value)}")
                    return None, option_value
                elif option_type == 'new_custom':
                    custom_path = input("Enter custom profile path or name: ").strip()
                    # If just a name without path, put it in home directory
                    if not os.path.sep in custom_path and not custom_path.startswith('~'):
                        custom_path = os.path.expanduser(f"~/{custom_path}")
                    else:
                        custom_path = os.path.expanduser(custom_path)

                    print(f"→ Using custom path: {custom_path}")
                    save_custom_profile(custom_path)
                    return None, custom_path
            else:
                print("Invalid choice! Please try again.")
        except ValueError:
            print("Please enter a valid number or profile name!")

def load_grades(input_file):
    """Load grades from CSV or Excel file"""
    print(f"📂 Loading {input_file}...")
    try:
        if input_file.endswith('.csv'):
            df = pd.read_csv(input_file)
        else:
            df = pd.read_excel(input_file)
    except Exception as e:
        print(f"❌ Error loading file: {e}")
        sys.exit(1)

    # Clean up the dataframe - remove empty rows
    df = df.dropna(how='all')

    # Check if this is a Moodle export format (has specific columns)
    if 'Last name' in df.columns and 'First name' in df.columns and 'Email address' in df.columns:
        print("✓ Detected Moodle export format")
        # Create a combined name column for compatibility
        df['Student Name'] = df['First name'] + ' ' + df['Last name']
        df['Email'] = df['Email address']

        # Filter out rows without email or with "Excused" grades
        df = df[df['Email address'].notna()]
        df = df[~df['Grade'].isin(['Excused', 'nan', ''])]
        df = df[df['Grade'].notna()]

        print(f"✓ Found {len(df)} students with valid grades")
    else:
        print(f"✓ Found {len(df)} entries")

    return df

def extract_student_data_from_page(driver):
    """Extract student data from Moodle page including emails and IDs"""
    js_code = """
    var students = {};

    // Method 1: Find all quickgrade inputs and work backwards
    var gradeInputs = document.querySelectorAll('input[name^="quickgrade_"]');

    console.log('Found ' + gradeInputs.length + ' grade input fields');

    for (var i = 0; i < gradeInputs.length; i++) {
        var input = gradeInputs[i];
        var userId = input.name.replace('quickgrade_', '');

        // Find the parent row
        var row = input.closest('tr');

        if (row) {
            // Try multiple selectors for email
            var email = '';
            var emailCell = row.querySelector('td.email') ||
                           row.querySelector('td.c2') ||
                           row.querySelector('td[class*="email"]');

            if (emailCell) {
                email = emailCell.textContent.trim();
            }

            // Try multiple selectors for name
            var name = '';
            var nameCell = row.querySelector('td.username') ||
                          row.querySelector('td.c1') ||
                          row.querySelector('td[class*="username"]');

            if (nameCell) {
                // Get all text content
                var fullText = nameCell.textContent.trim();
                // Clean up common prefixes
                name = fullText.replace(/Select\\s*/g, '')
                              .replace(/Picture of\\s*/g, '')
                              .replace(/^[A-Z]{2}/, ''); // Remove initials at start

                // If name has email in it, extract just the name part
                if (name.includes('@')) {
                    name = name.split('@')[0];
                }
            }

            // Get current grade value
            var currentGrade = input.value || '';

            // Store the student data
            students[userId] = {
                id: userId,
                email: email.toLowerCase(),
                name: name.trim(),
                current_grade: currentGrade,
                field_exists: true
            };

            console.log('Extracted student ' + userId + ': ' + name + ' (' + email + ')');
        }
    }

    // Method 2: If no emails found, try alternative approach
    if (Object.keys(students).length > 0) {
        var hasEmails = false;
        for (var id in students) {
            if (students[id].email) {
                hasEmails = true;
                break;
            }
        }

        if (!hasEmails) {
            console.log('No emails found, trying alternative extraction...');

            // Try to find emails from the full table
            var allCells = document.querySelectorAll('td');
            var emailPattern = /[a-z0-9._%+-]+@[a-z0-9.-]+\\.[a-z]{2,}/gi;

            for (var j = 0; j < allCells.length; j++) {
                var cellText = allCells[j].textContent;
                var emailMatches = cellText.match(emailPattern);

                if (emailMatches && emailMatches.length > 0) {
                    var parentRow = allCells[j].closest('tr');
                    if (parentRow) {
                        // Find if this row has a quickgrade input
                        var rowInput = parentRow.querySelector('input[name^="quickgrade_"]');
                        if (rowInput) {
                            var rowUserId = rowInput.name.replace('quickgrade_', '');
                            if (students[rowUserId]) {
                                students[rowUserId].email = emailMatches[0].toLowerCase();
                                console.log('Added email for ' + rowUserId + ': ' + emailMatches[0]);
                            }
                        }
                    }
                }
            }
        }
    }

    console.log('Total students extracted: ' + Object.keys(students).length);
    return students;
    """

    result = driver.execute_script(js_code)

    # Also check browser console for debugging
    try:
        logs = driver.get_log('browser')
        for log in logs:
            if 'console' in log['level'].lower():
                print(f"  Browser console: {log['message']}")
    except:
        pass  # Some configurations don't support browser logs

    return result

def map_grades_to_students(df, moodle_students):
    """Map CSV grades to Moodle students using email as primary key"""

    # Hardcoded mappings as fallback (when emails not available)
    NAME_TO_ID_FALLBACK = {
        "Vladislav Borisov": "4407",
        "Mukhammadamin Khatamov": "5254",
        "Dmitry Ershov": "5288",
        "Mekan Saryyev": "6425",
        "Aliia Khadeeva": "6428"
    }

    # Find columns - handle both formats
    email_col = None
    name_col = None
    grade_col = None
    feedback_col = None

    for col in df.columns:
        col_lower = col.lower()
        # Check for email columns (Email, Email address)
        if 'email' in col_lower:
            email_col = col
        # Check for name columns (Student Name, or use First/Last name)
        elif 'student name' in col_lower or (name_col is None and 'name' in col_lower):
            name_col = col
        # Check for grade columns
        elif 'grade' in col_lower or 'score' in col_lower:
            grade_col = col
        # Check for feedback columns
        elif 'feedback' in col_lower or 'comment' in col_lower:
            feedback_col = col

    if not grade_col:
        print("❌ Error: Need a column with 'Grade' or 'Score' in the CSV/Excel")
        sys.exit(1)

    if not email_col and not name_col:
        print("❌ Error: Need either 'Email' or 'Name' column in the CSV/Excel")
        sys.exit(1)

    # Check if Moodle students have emails
    has_emails_in_moodle = any(s.get('email', '') for s in moodle_students.values())

    # Build grades mapping
    matched_grades = {}
    unmatched = []
    multiple_matches = []

    print("\n📋 Matching students...")

    # If no emails in Moodle data, use fallback mappings
    if not has_emails_in_moodle:
        print("   Using hardcoded ID mappings (no emails found on page)...")

    for _, row in df.iterrows():
        csv_email = str(row[email_col]).lower().strip() if email_col and pd.notna(row[email_col]) else None
        csv_name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else None
        csv_grade = str(row[grade_col])
        csv_feedback = str(row[feedback_col]) if feedback_col and pd.notna(row[feedback_col]) else ""

        matches = []

        # If Moodle has no emails, use fallback mapping
        if not has_emails_in_moodle and csv_name:
            # Check hardcoded mappings
            for known_name, known_id in NAME_TO_ID_FALLBACK.items():
                if known_name.lower() in csv_name.lower() or csv_name.lower() in known_name.lower():
                    if known_id in moodle_students:
                        matches.append((known_id, moodle_students[known_id], 'fallback'))
                        break
        else:
            # Try to match by email first (most reliable)
            if csv_email and has_emails_in_moodle:
                for student_id, student_data in moodle_students.items():
                    if student_data.get('email', '') == csv_email:
                        matches.append((student_id, student_data, 'email'))
                        break  # Email should be unique, stop after first match

            # If no email match, try name (less reliable)
            if not matches and csv_name:
                for student_id, student_data in moodle_students.items():
                    moodle_name = student_data.get('name', '').lower()
                    if moodle_name and (csv_name.lower() in moodle_name or moodle_name in csv_name.lower()):
                        matches.append((student_id, student_data, 'name'))

        # Handle matches
        if len(matches) == 1:
            student_id, student_data, match_type = matches[0]
            matched_grades[student_id] = {
                'name': csv_name or student_data['name'],
                'email': csv_email or student_data['email'],
                'new_grade': csv_grade,
                'new_feedback': csv_feedback,
                'current_grade': student_data['current_grade'],
                'match_type': match_type
            }

            if match_type == 'email':
                match_icon = "📧"
            elif match_type == 'fallback':
                match_icon = "🔗"
            else:
                match_icon = "👤"
            print(f"  {match_icon} Matched: {csv_name} ({csv_email or 'no email'}) → Moodle ID {student_id}")

        elif len(matches) > 1:
            multiple_matches.append({
                'csv_name': csv_name,
                'csv_email': csv_email,
                'matches': [(m[0], m[1]['name'], m[1]['email']) for m in matches]
            })
            print(f"  ⚠️  Multiple matches for: {csv_name} ({csv_email})")

        else:
            unmatched.append({
                'name': csv_name,
                'email': csv_email,
                'grade': csv_grade
            })
            print(f"  ❌ No match found: {csv_name} ({csv_email})")

    return matched_grades, unmatched, multiple_matches

def display_analysis(matched_grades, unmatched, multiple_matches):
    """Display detailed analysis of grade matching"""

    empty_grades = []
    filled_grades = []

    for student_id, data in matched_grades.items():
        if data['current_grade'] and data['current_grade'].strip():
            filled_grades.append(data)
        else:
            empty_grades.append(data)

    print("\n" + "="*70)
    print("📊 GRADE MATCHING ANALYSIS")
    print("="*70)

    # Statistics
    print(f"\n📈 STATISTICS:")
    print(f"  Total in CSV: {len(matched_grades) + len(unmatched)}")
    print(f"  Successfully matched: {len(matched_grades)}")
    email_matches = sum(1 for d in matched_grades.values() if d.get('match_type') == 'email')
    name_matches = sum(1 for d in matched_grades.values() if d.get('match_type') == 'name')
    fallback_matches = sum(1 for d in matched_grades.values() if d.get('match_type') == 'fallback')

    if email_matches > 0:
        print(f"    - By email: {email_matches}")
    if name_matches > 0:
        print(f"    - By name: {name_matches}")
    if fallback_matches > 0:
        print(f"    - By hardcoded ID: {fallback_matches}")
    print(f"  Unmatched: {len(unmatched)}")
    print(f"  Multiple matches: {len(multiple_matches)}")

    # Empty grades (will be filled)
    if empty_grades:
        print(f"\n{Fore.GREEN}✓ STUDENTS WITHOUT GRADES (Will be filled):{Style.RESET_ALL}")
        print("-" * 50)
        for student in empty_grades:
            if student.get('match_type') == 'email':
                match_type = "📧"
            elif student.get('match_type') == 'fallback':
                match_type = "🔗"
            else:
                match_type = "👤"
            print(f"  {match_type} {student['name']} ({student['email']})")
            print(f"     → Grade: {Fore.GREEN}{student['new_grade']}{Style.RESET_ALL}")

    # Existing grades (need confirmation)
    if filled_grades:
        print(f"\n{Fore.YELLOW}⚠ STUDENTS WITH EXISTING GRADES:{Style.RESET_ALL}")
        print("-" * 50)
        for student in filled_grades:
            if student.get('match_type') == 'email':
                match_type = "📧"
            elif student.get('match_type') == 'fallback':
                match_type = "🔗"
            else:
                match_type = "👤"
            print(f"  {match_type} {student['name']} ({student['email']})")
            print(f"     Current: {Fore.RED}{student['current_grade']}{Style.RESET_ALL}")
            print(f"     New:     {Fore.GREEN}{student['new_grade']}{Style.RESET_ALL}")
            if student['current_grade'] != student['new_grade']:
                print(f"     {Fore.YELLOW}↳ Grade will change!{Style.RESET_ALL}")

    # Unmatched students
    if unmatched:
        print(f"\n{Fore.RED}✗ UNMATCHED STUDENTS (Cannot be graded):{Style.RESET_ALL}")
        print("-" * 50)
        for student in unmatched:
            print(f"  • {student['name']} ({student['email']}) - Grade: {student['grade']}")
        print(f"\n  {Fore.YELLOW}Tip: Check if these students are enrolled in the course{Style.RESET_ALL}")

    # Multiple matches warning
    if multiple_matches:
        print(f"\n{Fore.MAGENTA}⚠️  AMBIGUOUS MATCHES (Skipped):{Style.RESET_ALL}")
        print("-" * 50)
        for item in multiple_matches:
            print(f"  • {item['csv_name']} ({item['csv_email']}) matches:")
            for mid, mname, memail in item['matches']:
                print(f"    → ID: {mid}, {mname} ({memail})")

    return empty_grades, filled_grades

def get_user_choice(filled_grades, unmatched):
    """Get user choice for handling existing grades"""

    print("\n" + "="*70)
    print("❓ WHAT WOULD YOU LIKE TO DO?")
    print("="*70)

    if filled_grades:
        print(f"\n{len(filled_grades)} students already have grades.")
    if unmatched:
        print(f"{Fore.RED}{len(unmatched)} students could not be matched.{Style.RESET_ALL}")

    print("\nOptions:")
    print("  1. Fill ONLY empty grades (skip existing)")
    print("  2. OVERWRITE all grades (including existing)")
    print("  3. CANCEL operation")

    while True:
        choice = input("\nEnter your choice (1/2/3): ").strip()
        if choice == '1':
            return 'skip_existing'
        elif choice == '2':
            if filled_grades:
                confirm = input(f"\n{Fore.RED}⚠️  Are you sure you want to OVERWRITE {len(filled_grades)} existing grades? (yes/no): {Style.RESET_ALL}").strip().lower()
                if confirm == 'yes':
                    return 'overwrite'
                else:
                    print("Overwrite cancelled.")
            else:
                return 'overwrite'
        elif choice == '3':
            return 'cancel'
        else:
            print("Invalid choice. Please enter 1, 2, or 3.")

def inject_grades_smart(driver, matched_grades, mode):
    """Inject grades based on user choice"""

    # Prepare data for JavaScript
    grades_for_js = {}
    for student_id, data in matched_grades.items():
        grades_for_js[student_id] = {
            'grade': data['new_grade'],
            'feedback': data['new_feedback'],
            'current_grade': data['current_grade']
        }

    js_code = f"""
    var grades = {json.dumps(grades_for_js)};
    var mode = '{mode}';

    var stats = {{filled_new: 0, overwritten: 0, skipped: 0, errors: 0}};

    for (var id in grades) {{
        var gradeInput = document.querySelector('input[name="quickgrade_' + id + '"]');
        var feedbackTextarea = document.querySelector('textarea[name="quickgrade_comments_' + id + '"]');

        if (gradeInput) {{
            var hasExisting = grades[id].current_grade && grades[id].current_grade.trim();

            if (mode === 'skip_existing' && hasExisting) {{
                // Skip this student
                stats.skipped++;
                gradeInput.style.backgroundColor = '#FFF9C4'; // Light yellow
                gradeInput.title = 'Skipped - already has grade';
            }} else {{
                // Fill the grade
                gradeInput.value = grades[id].grade;

                if (hasExisting) {{
                    gradeInput.style.backgroundColor = '#FFAB91'; // Light orange for overwritten
                    gradeInput.style.border = '2px solid #FF6F00';
                    stats.overwritten++;
                }} else {{
                    gradeInput.style.backgroundColor = '#90EE90'; // Light green for new
                    gradeInput.style.border = '2px solid #4CAF50';
                    stats.filled_new++;
                }}

                if (feedbackTextarea && grades[id].feedback) {{
                    feedbackTextarea.value = grades[id].feedback;
                    feedbackTextarea.style.backgroundColor = gradeInput.style.backgroundColor;
                }}
            }}
        }} else {{
            stats.errors++;
        }}
    }}

    // Highlight save button
    var saveBtn = document.querySelector('input[value*="Save"], input[name="savechanges"], button[type="submit"]');
    if (saveBtn) {{
        saveBtn.style.backgroundColor = '#4CAF50';
        saveBtn.style.color = 'white';
        saveBtn.style.fontSize = '18px';
        saveBtn.style.padding = '10px';
        saveBtn.style.border = '3px solid #2E7D32';
        saveBtn.scrollIntoView({{behavior: 'smooth', block: 'center'}});
    }}

    return stats;
    """

    return driver.execute_script(js_code)

def setup_chrome_driver(profile_name=None, custom_path=None):
    """Setup Chrome driver with selected profile"""

    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)

    # Add arguments to bypass detection
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--no-sandbox')

    # Enable browser console logging
    options.set_capability('goog:loggingPrefs', {'browser': 'ALL'})

    if custom_path:
        options.add_argument(f"--user-data-dir={custom_path}")
        print(f"   Using custom profile path: {custom_path}")
    elif profile_name:
        system = platform.system()
        if system == "Linux":
            base_path = os.path.expanduser("~/.config/google-chrome")
        elif system == "Windows":
            base_path = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")
        elif system == "Darwin":
            base_path = os.path.expanduser("~/Library/Application Support/Google/Chrome")

        options.add_argument(f"--user-data-dir={base_path}")
        options.add_argument(f"--profile-directory={profile_name}")
        print(f"   Using Chrome profile: {profile_name}")
        print("   ⚠️  Make sure Chrome is closed if using this profile!")
    else:
        # Temporary profile
        temp_profile = os.path.expanduser("~/.moodle_temp_profile")
        options.add_argument(f"--user-data-dir={temp_profile}")
        print("   Using temporary profile")

    # Auto-download and setup ChromeDriver
    print("   Setting up ChromeDriver (auto-downloading if needed)...")
    service = Service(ChromeDriverManager().install())

    return webdriver.Chrome(service=service, options=options)

def main():
    # Parse arguments
    if len(sys.argv) < 2:
        print("Usage: python smart_grade_injector_v4_final.py <grades.csv/xlsx> [moodle_url]")
        print("\nCSV should have columns: Email, Student Name, Grade, Feedback")
        print("Email is used as the primary unique identifier")
        sys.exit(1)

    input_file = sys.argv[1]

    # Default Moodle URL (can be overridden)
    moodle_url = "https://moodle.innopolis.university/mod/assign/view.php?id=131683&action=grading"
    if len(sys.argv) > 2:
        moodle_url = sys.argv[2]

    print("\n" + "="*70)
    print("🎯 SMART MOODLE GRADE INJECTOR v4 FINAL")
    print("   Complete Edition with Profile Management")
    print("="*70)

    # Load grades from CSV
    df = load_grades(input_file)

    # Select Chrome profile
    profile_name, custom_path = select_chrome_profile()

    print(f"\n🌐 Target URL: {moodle_url}")

    if profile_name or custom_path:
        input("\nPress Enter to start (close Chrome if it's open)...")
    else:
        print("\n⚠️  You'll need to log in to Moodle manually")
        input("\nPress Enter to start...")

    # Setup Chrome
    print("\n🚀 Starting Chrome...")

    try:
        driver = setup_chrome_driver(profile_name, custom_path)
        print("✓ Chrome started successfully!")
    except Exception as e:
        print(f"\n❌ Error starting Chrome: {e}")

        if "user data directory" in str(e).lower():
            print("\n💡 SOLUTION: Close ALL Chrome windows and try again!")
            print("   Or select a different profile option.")

        sys.exit(1)

    try:
        # Navigate to Moodle
        print(f"\n📍 Navigating to Moodle...")
        driver.get(moodle_url)

        # Wait for page load
        print("\n⏳ Waiting for page to load...")
        if not (profile_name or custom_path):
            print("   (Please log in if prompted)")

        # Wait for grading table
        try:
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name^='quickgrade_']"))
            )
            print("✓ Grading page loaded!")
        except:
            print("\n❌ Timeout: Could not find grading table")
            print("   Make sure:")
            print("   - You're logged in")
            print("   - Quick grading is enabled")
            print("   - You're on the correct page")
            input("\nPress Enter to close browser...")
            driver.quit()
            sys.exit(1)

        # Wait for full page load
        time.sleep(2)

        # Extract student data from Moodle page
        print("\n🔍 Extracting student data from Moodle...")
        moodle_students = extract_student_data_from_page(driver)
        print(f"✓ Found {len(moodle_students)} students on Moodle page")

        # If no students found, try simpler extraction
        if len(moodle_students) == 0:
            print("⚠️  No students found with email extraction, trying name-only extraction...")
            js_simple = """
            var students = {};
            var inputs = document.querySelectorAll('input[name^="quickgrade_"]');
            for (var i = 0; i < inputs.length; i++) {
                var id = inputs[i].name.replace('quickgrade_', '');
                students[id] = {
                    id: id,
                    email: '',
                    name: 'Student_' + id,
                    current_grade: inputs[i].value || '',
                    field_exists: true
                };
            }
            return students;
            """
            moodle_students = driver.execute_script(js_simple)
            print(f"✓ Found {len(moodle_students)} grade input fields")

            if len(moodle_students) > 0:
                print("\n⚠️  NOTE: Could not extract emails from page.")
                print("   Will use the hardcoded ID mappings instead.")

        # Map CSV grades to Moodle students
        matched_grades, unmatched, multiple_matches = map_grades_to_students(df, moodle_students)

        # Display analysis
        empty_grades, filled_grades = display_analysis(matched_grades, unmatched, multiple_matches)

        if not matched_grades:
            print(f"\n{Fore.RED}❌ No students could be matched! Check your CSV file.{Style.RESET_ALL}")
            driver.quit()
            sys.exit(1)

        # Get user choice
        choice = get_user_choice(filled_grades, unmatched)

        if choice == 'cancel':
            print("\n❌ Operation cancelled by user")
            driver.quit()
            sys.exit(0)

        # Inject grades based on choice
        print(f"\n💉 Injecting grades (mode: {choice})...")
        stats = inject_grades_smart(driver, matched_grades, choice)

        # Display results
        print("\n" + "="*70)
        print("✅ INJECTION COMPLETE!")
        print("="*70)
        if stats['filled_new'] > 0:
            print(f"  {Fore.GREEN}✓ New grades filled: {stats['filled_new']}{Style.RESET_ALL}")
        if stats['overwritten'] > 0:
            print(f"  {Fore.YELLOW}✓ Grades overwritten: {stats['overwritten']}{Style.RESET_ALL}")
        if stats['skipped'] > 0:
            print(f"  {Fore.CYAN}○ Grades skipped: {stats['skipped']}{Style.RESET_ALL}")
        if stats['errors'] > 0:
            print(f"  {Fore.RED}✗ Errors: {stats['errors']}{Style.RESET_ALL}")

        print("\n📌 COLOR GUIDE:")
        print(f"  {Fore.GREEN}Green fields{Style.RESET_ALL} = New grades added")
        print(f"  {Fore.YELLOW}Orange fields{Style.RESET_ALL} = Existing grades overwritten")
        print(f"  {Fore.CYAN}Yellow fields{Style.RESET_ALL} = Skipped (already had grades)")

        print("\n" + "="*70)
        print("📌 NEXT STEPS:")
        print("="*70)
        print("1. Review the colored grade fields")
        print("2. Click the green 'Save all quick grading changes' button")
        print("3. Wait for Moodle to confirm the save")

        input("\nPress Enter after you've saved the grades...")

    except KeyboardInterrupt:
        print("\n\n❌ Interrupted by user")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        input("\nPress Enter to close browser...")
    finally:
        driver.quit()
        print("\n👍 Browser closed. Done!")

if __name__ == "__main__":
    main()