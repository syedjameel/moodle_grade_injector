# Moodle Grade Injector

A simple script for bulk grading in Moodle's quick grading interface. This tool helps instructors efficiently upload grades from CSV or Excel files directly into Moodle, saving significant time on manual grade entry.

## Features

- **Bulk Grade Upload**: Import grades from CSV or Excel files directly into Moodle
- **Student Matching**: Automatically matches students using email addresses or names
- **Chrome Profile Support**: Save login sessions across multiple grading sessions
- **Visual Feedback**: Color-coded interface shows which grades are new, updated, or skipped
- **Safety Features**: Preview changes before applying, with options to skip existing grades
- **Auto dependency Installation**: Automatically installs required Python packages

## Prerequisites

- Python 3.6 or higher
- Google Chrome browser
- Active Moodle account with grading permissions
- Linux

## Installation

1. clone and go to this repository:
```bash
git clone https://github.com/syedjameel/moodle_grade_injector.git
cd moodle_grade_injector
```

2. The script will automatically install required dependencies on first run. Alternatively, install manually:
```bash
pip install pandas selenium webdriver-manager colorama openpyxl
```

## Usage

### Basic Command

```bash
python moodle_grade_injector.py <grades_file> <moodle_url>
```

### Examples

```bash
# Using a CSV file (recommendeed)
python moodle_grade_injector.py homework1_grades_sample.csv "https://moodle.innopolis.university/mod/assign/view.php?id=131686&action=grading"

### Example CSV Structure:
```csv
Last name,First name,Groups,Student ID,Email address,Grade,Feedback
Ivanov,Ivan,B24-CSE-07,1234,ivan.xyz@innopolis.university,100,Very Good
Vladislav,Vlad,B24-CSE-07,1234,vlad.xyz@innopolis.university,100,Very Good
```

## Step-by-Step Workflow

### 1. Prepare Your Grades
- Create a CSV/Excel file with student emails and grades
- Ensure email addresses match those in Moodle
- Save the file in the same directory as the script (or provide full path)

### 2. Get the Moodle URL
- Navigate to your course in Moodle
- Go to the assignment/quiz you want to grade
- Click on "Submissions" or similar
- Enable "Quick grading" if not already enabled
- Copy the URL from your browser

### 3. Run the Script
```bash
python moodle_grade_injector.py grades.csv "YOUR_MOODLE_URL"
```

### 4. Chrome Profile Selection (Recommened to use custom profile)
When prompted, choose how to handle Chrome login:

- **0** - Create temporary profile (requires login each time)
- **1-N** - Use existing Chrome profile (stays logged in)
- **Custom** - Use a dedicated profile for grading

**Recommended**: Create a custom profile called "moodle-grading" for persistent login

### 5. Login to Moodle (if needed)
- If using a new profile, manually log in to Moodle when the browser opens
- The script waits for you to complete the login

### 6. Review Matching Results
The script will display:
- Successfully matched students (by email or name)
- Students with existing grades
- Unmatched students (not found in Moodle)

### 7. Choose Action
Select how to handle the grades:
1. **Fill ONLY empty grades** - Skip students who already have grades
2. **OVERWRITE all grades** - Update all grades including existing ones
3. **CANCEL** - Exit without making changes

### 8. Review and Save
- The script highlights modified fields:
  - **Green** = New grades added
  - **Orange** = Existing grades overwritten
  - **Yellow** = Skipped (already had grades)
- Click the "Save all quick grading changes" button in Moodle
- Press Enter in the terminal after saving


## Troubleshooting

### "Chrome driver not found" Error
The script uses `webdriver-manager` to auto-download ChromeDriver. If this fails:
1. Ensure Chrome browser is installed
2. Check your internet connection
3. Try running with administrator/sudo privileges

### "User data directory already in use" Error
This occurs when Chrome is already open with the same profile:
1. Close all Chrome windows
2. Run the script again
3. Or choose a different profile option

### Students Not Matching
If students aren't being matched:
1. Verify email addresses in CSV match exactly with Moodle
2. Check for typos or extra spaces
3. Ensure students are enrolled in the course
4. Try using student names as backup matching method

### Page Not Loading
If the grading page doesn't load:
1. Verify the URL is correct
2. Ensure you have grading permissions
3. Check that "Quick grading" is enabled in Moodle
4. Increase the timeout in the script if needed

## Safety Features

- **Preview Mode**: Always shows what will be changed before applying
- **Confirmation Required**: Asks for confirmation before overwriting existing grades
- **Visual Indicators**: Color coding makes it clear what's being modified
- **No Auto-Save**: You must manually click save in Moodle after review


## Common Use Cases

### Homework 1 Grading
```bash
python moodle_grade_injector.py homework1_grades_sample.csv "https://moodle.innopolis.university/mod/assign/view.php?id=131686&action=grading"
```

## Security Notes

- The script never stores your Moodle credentials
- Chrome profiles are stored locally on your computer
- No data is sent to external servers
- All processing happens locally on your machine