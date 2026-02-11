from flask import Flask, render_template, request, send_file
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime
import pandas as pd
import os
from math import radians, cos, sin, asin, sqrt
from openpyxl.styles import Font, PatternFill
from openpyxl.drawing.image import Image as OpenpyxlImage

app = Flask(__name__)

# CONFIGURATION
CAMPUS_LAT = -26.099059
CAMPUS_LON = 28.0538272
MAX_DISTANCE_KM = 0.2 
START_DATE = datetime(2026, 2, 13)

def get_db_url():
    """Handles Render's database URL formatting requirement."""
    url = os.environ.get('DATABASE_URL')
    if url and url.startswith("postgres://"):
        return url.replace("postgres://", "postgresql://", 1)
    return url

def get_db_connection():
    return psycopg2.connect(get_db_url())

def init_db():
    """Ensures the table exists before any requests are made."""
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS check_ins (
            id SERIAL PRIMARY KEY,
            first_name TEXT NOT NULL,
            middle_name TEXT,
            last_name TEXT NOT NULL,
            id_passport TEXT NOT NULL,
            email TEXT NOT NULL,
            phone TEXT NOT NULL,
            ethnicity TEXT,
            gender TEXT,
            course TEXT,
            week_number INTEGER,
            check_in_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    cur.close()
    conn.close()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    user_lat = request.form.get('latitude')
    user_lon = request.form.get('longitude')
    
    if not user_lat or not user_lon:
        return "<h1>Location Required</h1><p>Please enable GPS.</p>"

    # Distance calculation logic here...
    
    data = (
        request.form.get('first_name'), request.form.get('middle_name'),
        request.form.get('last_name'), request.form.get('id_passport'),
        request.form.get('email'), request.form.get('phone'),
        request.form.get('ethnicity'), request.form.get('gender'),
        request.form.get('course'), 1 # Default week for testing
    )

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute('''INSERT INTO check_ins 
        (first_name, middle_name, last_name, id_passport, email, phone, ethnicity, gender, course, week_number)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)''', data)
    conn.commit()
    cur.close()
    conn.close()
    return render_template('success.html', name=data[0])

@app.route('/admin')
def admin():
    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('SELECT * FROM check_ins ORDER BY check_in_time DESC')
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('admin.html', rows=rows)

# ... (Include your download route from before)

if __name__ == '__main__':
    # Initialize DB every time the app starts
    try:
        init_db()
    except Exception as e:
        print(f"Startup Error: {e}")
        
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
