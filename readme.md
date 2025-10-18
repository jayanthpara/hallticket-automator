# Hallticket Automator

Generate individual hall tickets for students easily! This project automates the creation of student hall tickets in PDF format from a given list and template—saving time and reducing manual errors[web:8][web:6].

## Features

- Accepts student data in bulk (CSV/Excel)
- Uses custom templates for hall tickets
- Generates personalized hall ticket PDFs for each entry
- Simple setup and clear workflow

## Requirements

- Python 3.7+
- pandas, fpdf (or any other PDF library used)

## Setup

1. Clone the repository:
git clone https://github.com/jayanthpara/hallticket-automator.git

cd hallticket-automator

2. Install dependencies:
pip install -r requirements.txt

3. Prepare your student list (`students.csv`) and hall ticket template (`template.html` or as required).

## Usage
python generate_halltickets.py --list students.csv --template template.html --output ./halltickets
- This command will generate PDFs and save them in the specified output directory.

## Documentation

### Input Formats

- Student List: Should be in CSV format with columns such as Name, Roll Number, Exam Date, etc.
- Template: HTML or text file with placeholders (e.g., {{Name}}, {{RollNumber}}).

### Output

- Individual PDF files for each student, named using their roll number or name.

### Extending

- Change/add placeholders in the template as needed.
- Supports batch generation; easily integrate or automate for different exams.

## Contribution

Feel free to fork this repo, raise issues or contribute through pull requests!

