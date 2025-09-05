from flask import Flask, render_template, request, flash, redirect, session, send_file, url_for
import pandas as pd
from io import BytesIO
import json
from datetime import datetime
from collections import defaultdict
from ga_model import initialize_population, calculate_fitness, genetic_algorithm, check_constraints
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
# Set a strong secret key for session management
app.config['SECRET_KEY'] = 'your-secure-secret-key-here'  # Change this in production
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Allowed file extensions
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """Render the main page."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file uploads and generate timetable."""
    try:
        # Check if files were uploaded
        if 'data_exam' not in request.files or 'data_invigilator' not in request.files or 'contoh_jadual' not in request.files:
            flash("All files must be uploaded.", "error")
            return redirect(url_for('index'))

        data_exam_file = request.files['data_exam']
        data_invigilator_file = request.files['data_invigilator']
        contoh_jadual_file = request.files['contoh_jadual']

        # Validate file extensions
        for file in [data_exam_file, data_invigilator_file, contoh_jadual_file]:
            if not allowed_file(file.filename):
                flash("Invalid file type. Please upload Excel files only.", "error")
                return redirect(url_for('index'))

        # Load and validate exam data
        try:
            data_exam = pd.read_excel(data_exam_file, header=0)
            data_exam.columns = [
                'Tarikh', 'Hari', 'Waktu', 'Kod Kursus', 'Kod Setara', 
                'Masa Mula', 'Masa Tamat', 'Jumlah Pelajar'
            ]
        except Exception as e:
            flash("Error reading exam data file. Please check the format.", "error")
            return redirect(url_for('index'))

        # Load and validate invigilator data
        try:
            invigilator_sheets = pd.ExcelFile(data_invigilator_file).sheet_names
            required_sheets = ['STAFF(S)', 'LECTERUR(L & K)']
            
            if not all(sheet in invigilator_sheets for sheet in required_sheets):
                flash("Missing required worksheets in invigilator file.", "error")
                return redirect(url_for('index'))

            staff_df = pd.read_excel(data_invigilator_file, sheet_name='STAFF(S)')
            lecturer_df = pd.read_excel(data_invigilator_file, sheet_name='LECTERUR(L & K)')
            
            staff_df.columns = ['Bil', 'No. Pekerja', 'Nama', 'Jawatan', 'Jantina']
            lecturer_df.columns = ['Bil', 'No. Pekerja', 'Nama', 'Jawatan', 'Jantina']
        except Exception as e:
            flash("Error reading invigilator data file. Please check the format.", "error")
            return redirect(url_for('index'))

        # Load and validate sample schedule
        try:
            contoh_jadual = pd.read_excel(contoh_jadual_file, header=0)
            if len(contoh_jadual.columns) == 10:
                contoh_jadual.columns = [
                    'Tarikh', 'Masa', 'Kod Kursus', 'Kod Program', 'Kumpulan', 
                    'Bilangan Calon', 'Pensyarah', 'Tempat', 'Pengawas', 'Unused'
                ]
            else:
                flash("Invalid sample schedule format.", "error")
                return redirect(url_for('index'))
        except Exception as e:
            flash("Error reading sample schedule file. Please check the format.", "error")
            return redirect(url_for('index'))

        # Process data and generate timetable
        try:
            # Clean exam data
            data_exam['Tarikh'] = pd.to_datetime(data_exam['Tarikh'], errors='coerce')
            data_exam.dropna(subset=['Tarikh', 'Kod Kursus'], inplace=True)

            # Clean sample schedule
            contoh_jadual.drop(columns=['Pengawas', 'Unused'], inplace=True, errors='ignore')
            contoh_jadual['Tarikh'] = pd.to_datetime(contoh_jadual['Tarikh'], errors='coerce')
            contoh_jadual.dropna(subset=['Tarikh', 'Kod Kursus'], inplace=True)

            # Combine invigilator data
            data_invigilator = pd.concat([staff_df, lecturer_df], ignore_index=True).drop(columns=['Bil'])
            data_invigilator.dropna(subset=['Nama'], inplace=True)
            data_invigilator['Nama'] = data_invigilator['Nama'].str.lower().str.strip()

            # Generate timetable with violations check
            population = initialize_population(data_exam, data_invigilator, contoh_jadual, population_size=100)
            best_schedule, final_fitness, violations = genetic_algorithm(population, data_exam, data_invigilator)
            
            # Format schedule for display
            formatted_schedule = format_schedule(best_schedule)
            
            # Store in session
            session['current_schedule'] = formatted_schedule
            session['violations'] = violations
            
            flash("Timetable generated successfully! You can now view the dashboard.", "success")
            
            return render_template(
                'index.html', 
                schedule=formatted_schedule, 
                fitness=final_fitness,
                violations=violations
            )

        except Exception as e:
            flash(f"Error generating timetable: {str(e)}", "error")
            return redirect(url_for('index'))

    except Exception as e:
        flash(f"An unexpected error occurred: {str(e)}", "error")
        return redirect(url_for('index'))

@app.route('/dashboard')
def dashboard():
    """Render the dashboard with statistics."""
    try:
        # Get schedule and violations from session
        schedule = session.get('current_schedule')
        violations = session.get('violations', {})

        # Initialize stats
        invigilator_stats = []
        exam_details = []
        overview_stats = {
            'total_exams': 0,
            'total_lecturers': 0,
            'total_staff': 0
        }

        if schedule:
            # Calculate statistics only if schedule exists
            invigilator_stats = calculate_invigilator_stats(schedule)
            exam_details = prepare_exam_details(schedule)
            overview_stats = calculate_overview_stats(schedule)

            return render_template(
                'dashboard.html',
                invigilator_stats=invigilator_stats,
                exam_details=exam_details,
                violations=violations,
                total_exams=overview_stats['total_exams'],
                total_lecturers=overview_stats['total_lecturers'],
                total_staff=overview_stats['total_staff']
            )
        else:
            # If no schedule in session, redirect with message
            flash("Please generate a timetable first before accessing the dashboard.", "warning")
            return redirect(url_for('index'))

    except Exception as e:
        # Log the error for debugging
        print(f"Dashboard error: {str(e)}")
        flash(f"An error occurred while loading the dashboard. Please try regenerating the timetable.", "error")
        return redirect(url_for('index'))

@app.route('/export', methods=['POST'])
def export_timetable():
    """Export timetable to Excel file."""
    try:
        schedule_data = request.form.get('schedule_data')
        if not schedule_data:
            flash("No schedule data available for export", "error")
            return redirect(url_for('index'))

        try:
            schedule = json.loads(schedule_data)
        except json.JSONDecodeError as e:
            flash(f"Error parsing schedule data: {str(e)}", "error")
            return redirect(url_for('index'))

        # Create DataFrame from schedule
        df = pd.DataFrame(schedule)
        
        # Ensure all required columns are present
        required_columns = [
            'Date', 'Time', 'Day', 'Course Code', 'Venue',
            'Lecturer(s)', 'Number of Students', 'Invigilator(s)'
        ]
        
        # Reorder columns if they exist
        existing_columns = [col for col in required_columns if col in df.columns]
        df = df[existing_columns]

        # Format date
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')

        # Create Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Timetable', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Timetable']

            # Format headers
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#3498db',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            # Apply formats
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # Adjust column widths
            for idx, col in enumerate(df.columns):
                max_length = max(
                    df[col].astype(str).apply(len).max(),
                    len(col)
                ) + 2
                worksheet.set_column(idx, idx, max_length)

        output.seek(0)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"examination_timetable_{timestamp}.xlsx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        flash(f"Error exporting timetable: {str(e)}", "error")
        return redirect(url_for('index'))

def calculate_invigilator_stats(schedule):
    """Calculate detailed statistics for each invigilator's sessions."""
    stats = []
    invigilator_session_count = defaultdict(int)
    
    # First pass to count total sessions per invigilator
    for exam in schedule:
        invigilators = exam['Invigilator(s)'].split(', ') if isinstance(exam['Invigilator(s)'], str) else exam['Invigilator(s)']
        for invigilator in invigilators:
            invigilator_session_count[invigilator] += 1
    
    # Second pass to create detailed session information
    for exam in schedule:
        invigilators = exam['Invigilator(s)'].split(', ') if isinstance(exam['Invigilator(s)'], str) else exam['Invigilator(s)']
        for invigilator in invigilators:
            stats.append({
                'name': invigilator,
                'date': exam['Date'],
                'day': exam['Day'],
                'time': exam['Time'],
                'course_code': exam['Course Code'],
                'venue': exam['Venue'],
                'total_sessions': invigilator_session_count[invigilator]
            })
    
    return sorted(stats, key=lambda x: (x['name'], x['date']))

def prepare_exam_details(schedule):
    """Prepare detailed exam information."""
    return [
        {
            'course_code': exam['Course Code'],
            'date': exam['Date'],
            'venue': exam['Venue'],
            'students': exam.get('Number of Students', ''),
            'invigilators': exam['Invigilator(s)'].split(', ') if isinstance(exam['Invigilator(s)'], str) else exam['Invigilator(s)']
        }
        for exam in schedule
    ]

def calculate_overview_stats(schedule):
    """Calculate overview statistics including proper staff count."""
    # Get all invigilators
    all_invigilators = [
        invigilator.strip()
        for exam in schedule
        for invigilator in (exam['Invigilator(s)'].split(', ') if isinstance(exam['Invigilator(s)'], str) else exam['Invigilator(s)'])
    ]
    
    # Count total staff (those marked with (S))
    total_staff = len(set(
        invigilator 
        for invigilator in all_invigilators 
        if '(S)' in invigilator
    ))
    
    # Count total lecturers (those marked with (L) or (K))
    total_lecturers = len(set(
        invigilator 
        for invigilator in all_invigilators 
        if '(L)' in invigilator or '(K)' in invigilator
    ))
    
    return {
        'total_exams': len(schedule),
        'total_lecturers': total_lecturers,
        'total_staff': total_staff
    }

def format_schedule(schedule):
    """Format schedule for display."""
    return [
        {
            "Course Code": exam['Kod Kursus'],
            "Date": exam['Tarikh'].strftime('%Y-%m-%d') if isinstance(exam['Tarikh'], pd.Timestamp) else exam['Tarikh'],
            "Day": exam['Hari'],
            "Time": exam['Masa'],
            "Venue": exam['Tempat'],
            "Lecturer(s)": exam['Lecturer'],
            "Invigilator(s)": ', '.join(exam['Invigilators']) if isinstance(exam['Invigilators'], list) else exam['Invigilators'],
            "Number of Students": exam['Bilangan Pelajar']
        }
        for exam in schedule
    ]

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
