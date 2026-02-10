from flask import Flask, render_template, request, send_file
import sqlite3
from datetime import datetime
import pandas as pd
import os
from math import radians, cos, sin, asin, sqrt
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)

# CONFIGURATION
CAMPUS_LAT = -26.099059
CAMPUS_LON = 28.0538272
MAX_DISTANCE_KM = 0.2  # 200 meter radius
START_DATE = datetime(2026, 2, 13)

def calculate_distance(lat1, lon1):
    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, CAMPUS_LON, CAMPUS_LAT])
    dlon, dlat = lon2 - lon1, lat2 - lat1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    return 2 * asin(sqrt(a)) * 6371

def get_db_connection():
    conn = sqlite3.connect('database.db')
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    conn.execute('''CREATE TABLE IF NOT EXISTS check_ins (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        first_name TEXT, middle_name TEXT, last_name TEXT,
        id_passport TEXT, email TEXT, phone TEXT,
        ethnicity TEXT, gender TEXT, course TEXT,
        week_number INTEGER,
        check_in_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')
    conn.commit()
    conn.close()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    user_lat = request.form.get('latitude')
    user_lon = request.form.get('longitude')
    
    if not user_lat or not user_lon:
        return "<h1>Location Error</h1><p>GPS is required to check in.</p>"

    if calculate_distance(float(user_lat), float(user_lon)) > MAX_DISTANCE_KM:
        return "<h1>Access Denied</h1><p>You must be at 85 Grayston Drive to check in.</p>"

    today = datetime.now()
    week_num = max(1, ((today - START_DATE).days // 7) + 1)
    
    data = (
        request.form.get('first_name'), request.form.get('middle_name'),
        request.form.get('last_name'), request.form.get('id_passport'),
        request.form.get('email'), request.form.get('phone'),
        request.form.get('ethnicity'), request.form.get('gender'),
        request.form.get('course'), week_num
    )

    conn = get_db_connection()
    conn.execute('''INSERT INTO check_ins 
        (first_name, middle_name, last_name, id_passport, email, phone, ethnicity, gender, course, week_number) 
        VALUES (?,?,?,?,?,?,?,?,?,?)''', data)
    conn.commit()
    conn.close()
    return f"<h1>Success!</h1><p>Thank you, {data[0]}. Attendance recorded.</p><a href='/'>Back</a>"

@app.route('/admin')
def admin():
    conn = get_db_connection()
    rows = conn.execute('SELECT * FROM check_ins ORDER BY check_in_time DESC').fetchall()
    conn.close()
    return render_template('admin.html', rows=rows)

@app.route('/download/<string:course>/<int:week>')
def download(course, week):
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM check_ins WHERE week_number = ? AND course = ?", 
                           conn, params=(week, course))
    conn.close()
    
    if df.empty:
        return f"<h1>No data</h1><p>No {course} check-ins for Week {week}.</p>"

    df.columns = ['ID', 'First Name', 'Middle Name', 'Last Name', 'ID/Passport', 'Email', 'Phone', 'Ethnicity', 'Gender', 'Course', 'Week', 'Time']
    os.makedirs('exports', exist_ok=True)
    filename = f"exports/SouthernLabs_{course}_Week_{week}.xlsx"
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Attendance')
        ws = writer.sheets['Attendance']
        fill = PatternFill(start_color='003366', end_color='003366', fill_type='solid')
        font = Font(color='FFFFFF', bold=True)
        for cell in ws[1]:
            cell.fill, cell.font = fill, font
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 20
            
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    if not os.path.exists('database.db'): 
        init_db()
    # Port is handled by Render's environment
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
