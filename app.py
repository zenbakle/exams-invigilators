from pathlib import Path
from io import BytesIO

import pandas as pd
from flask import Flask, render_template, request, send_file, make_response, url_for
from weasyprint import HTML

app = Flask(__name__)
DATA_DIR = Path(__file__).resolve().parent

ASSIGNMENTS_FILE = DATA_DIR / "invigilation_assignments1.xlsx"
STAFF_FILE = DATA_DIR / "staff_schedules1.xlsx"

# Load data once at startup
assignments_df = pd.read_excel(ASSIGNMENTS_FILE)
staff_df = pd.read_excel(STAFF_FILE)

assignments_df['date'] = pd.to_datetime(
    assignments_df['date'], errors='coerce')
if 'department' not in assignments_df.columns:
    assignments_df['department'] = assignments_df.get('staff_dept', '')

# Normalize the staff schedule department column to department for the website output
staff_df['department'] = staff_df.get('department', '')
staff_df['staff_email'] = staff_df['staff_email'].astype(
    str).str.strip().str.lower()
staff_df['date'] = pd.to_datetime(staff_df['date'], errors='coerce')

unique_departments = sorted(assignments_df['department'].dropna().unique())

OUTPUT_COLUMNS_assignments = ['date', 'time', 'course_code',
                              'course_title', 'venue', 'department', 'invigilators']
OUTPUT_COLUMNS_staff = ['date', 'time', 'course_code',
                        'course_title', 'venue', 'department']


def format_table_rows(df):
    rows = []
    for _, row in df.iterrows():
        rows.append({
            'date': row['date'].strftime('%d-%m-%Y') if pd.notna(row['date']) else '',
            'time': str(row.get('time', '')).strip(),
            'course_code': str(row.get('course_code', '')).strip(),
            'course_title': str(row.get('course_title', '')).strip(),
            'venue': str(row.get('venue', '')).strip(),
            'department': str(row.get('department', '')).strip(),
            'invigilators': str(row.get('invigilators', '')).strip(),
        })
    return rows


@app.route('/')
def home():
    return render_template('base.html')


@app.route('/assignments', methods=['GET'])
def assignments():
    selected_department = request.args.get('department', '').strip().lower()
    filtered_rows = []
    if selected_department:
        filtered_df = assignments_df[assignments_df['department'].str.strip(
        ).str.lower() == selected_department]
        filtered_df = filtered_df[OUTPUT_COLUMNS_assignments]
        filtered_rows = format_table_rows(filtered_df)

    return render_template(
        'assignments.html',
        departments=unique_departments,
        selected_department=selected_department,
        rows=filtered_rows,
    )


@app.route('/assignments/pdf', methods=['GET'])
def assignments_pdf():
    selected_department = request.args.get('department', '').strip().lower()
    if not selected_department:
        return "Please provide a department to download the PDF.", 400

    filtered_df = assignments_df[assignments_df['department'].str.strip(
    ).str.lower() == selected_department]
    filtered_df = filtered_df[OUTPUT_COLUMNS_assignments]
    rows = format_table_rows(filtered_df)
    title = f"Invigilation Assignments - {selected_department.title()}"

    html = render_template('pdf_table.html', title=title,
                           columns=OUTPUT_COLUMNS_assignments, rows=rows)
    pdf = HTML(string=html).write_pdf(stylesheets=None)

    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers[
        'Content-Disposition'] = f'attachment; filename=assignments_{selected_department}.pdf'
    return response


@app.route('/staff-schedule', methods=['GET'])
def staff_schedule():
    email = request.args.get('email', '').strip().lower()
    filtered_rows = []
    if email:
        filtered_df = staff_df[staff_df['staff_email'] == email]
        if not filtered_df.empty:
            filtered_df = filtered_df[OUTPUT_COLUMNS_staff]
            filtered_rows = format_table_rows(filtered_df)

    return render_template(
        'staff_schedule.html',
        email=email,
        rows=filtered_rows,
    )


@app.route('/staff-schedule/pdf', methods=['GET'])
def staff_schedule_pdf():
    email = request.args.get('email', '').strip().lower()
    if not email:
        return "Please provide an email to download the PDF.", 400

    filtered_df = staff_df[staff_df['staff_email'] == email]
    filtered_df = filtered_df[OUTPUT_COLUMNS_staff]
    rows = format_table_rows(filtered_df)
    title = f"Staff Schedule - {email}"

    html = render_template('pdf_table.html', title=title,
                           columns=OUTPUT_COLUMNS_staff, rows=rows)
    pdf = HTML(string=html).write_pdf(stylesheets=None)

    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers[
        'Content-Disposition'] = f'attachment; filename=staff_schedule_{email}.pdf'
    return response


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
