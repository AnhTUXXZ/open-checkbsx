import sys
import os
import time
import sqlite3
import threading
import socket
import io
import re
import base64
from datetime import datetime, timedelta

from flask import Flask, request, jsonify, render_template
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import openpyxl

import cv2
import numpy as np

# ==========================================
# 1. CẤU HÌNH HỆ THỐNG
# ==========================================

DATA_DIR = 'data'
TEMP_IMAGES_DIR = 'temp_images'
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(TEMP_IMAGES_DIR, exist_ok=True)

DB_PARKING = os.path.join(DATA_DIR, 'parking.db')

# Tự động nhận diện môi trường để cấu hình đường dẫn Tesseract OCR
IS_RENDER = os.environ.get('RENDER') is not None
if not IS_RENDER:
    pytesseract.pytesseract.tesseract_cmd = os.path.join(DATA_DIR, 'Tesseract-OCR', 'tesseract.exe')
else:
    pytesseract.pytesseract.tesseract_cmd = 'tesseract'

app = Flask(__name__)
app.secret_key = 'parking_admin_secret_key'
sheet_lock = threading.Lock()

EXCEL_FILE = 'thong_ke_bai_xe.xlsx'

def init_db():
    conn = sqlite3.connect(DB_PARKING)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS parking_logs
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                 plate TEXT, 
                 time_in DATETIME, 
                 time_out DATETIME, 
                 status TEXT)''')
    conn.commit()
    conn.close()

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "LỊCH SỬ RA VÀO"
        ws.append(["ID", "Biển Số", "Thời Gian Vào", "Thời Gian Ra", "Trạng Thái"])
        wb.save(EXCEL_FILE)

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

# Khởi tạo cơ sở dữ liệu và file excel ngay khi chạy app
init_db()
init_excel()

# ==========================================
# 2. FLASK ROUTES & XỬ LÝ OCR (WEB API)
# ==========================================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/cam')
def camera_scan():
    return render_template('cam.html')

@app.route('/admin')
def admin_dashboard():
    return render_template('admin.html')

@app.route('/ping', methods=['POST'])
def ping():
    return jsonify({'status': 'ok'})

@app.route('/upload-ocr', methods=['POST'])
def process_image():
    if 'image' not in request.files:
        return jsonify({'error': 'Không tìm thấy ảnh'}), 400
    
    file = request.files['image']
    try:
        in_memory_file = io.BytesIO(file.read())
        file_bytes = np.frombuffer(in_memory_file.getvalue(), dtype=np.uint8)
        img = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)

        # Tiền xử lý ảnh chuyên sâu
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        
        sharpen_kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
        sharpened = cv2.filter2D(gray, -1, sharpen_kernel)

        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        clahe_img = clahe.apply(sharpened)

        _, thresh = cv2.threshold(cv2.GaussianBlur(clahe_img, (5, 5), 0), 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # Đóng chữ (Nối các nét bị đứt do mờ)
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3,3))
        morph = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
        
        # MỚI: Mở chữ (Tẩy các vết ốc vít, chấm đen nhỏ để không đọc nhầm 0 thành 9)
        kernel_small = cv2.getStructuringElement(cv2.MORPH_RECT, (2,2))
        open_morph = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel_small)

        # Đưa thêm open_morph vào đầu mảng ưu tiên để AI quét trước
        filters = [open_morph, morph, thresh, clahe_img, sharpened, gray]
        psms = [11, 7, 6, 3, 12] 
        
        plate_text = ""
        best_d = None
        image_base64 = ""

        for f_img in filters:
            for psm in psms:
                custom_config = f'--oem 3 --psm {psm}'
                raw_text = pytesseract.image_to_string(f_img, config=custom_config)
                
                clean_text = ''.join(e for e in raw_text if e.isalnum()).upper()
                matches = re.findall(r'[1-9]\d[A-Z][A-Z0-9]?\d{4,5}', clean_text)
                
                if matches:
                    plate_text = max(matches, key=len)
                    best_d = pytesseract.image_to_data(f_img, config=custom_config, output_type=pytesseract.Output.DICT)
                    break 
            if plate_text:
                break 

        if plate_text:
            try:
                box_x_min, box_y_min, box_x_max, box_y_max = 9999, 9999, 0, 0
                for i in range(len(best_d['text'])):
                    word = ''.join(e for e in best_d['text'][i] if e.isalnum()).upper()
                    if word and (word in plate_text):
                        x = int(best_d['left'][i] / 2)
                        y = int(best_d['top'][i] / 2)
                        w = int(best_d['width'][i] / 2)
                        h = int(best_d['height'][i] / 2)
                        
                        box_x_min = min(box_x_min, x)
                        box_y_min = min(box_y_min, y)
                        box_x_max = max(box_x_max, x + w)
                        box_y_max = max(box_y_max, y + h)

                if box_x_min != 9999:
                    cv2.rectangle(img, (box_x_min - 15, box_y_min - 15), (box_x_max + 15, box_y_max + 15), (0, 255, 0), 3)
                    cv2.putText(img, plate_text, (box_x_min - 15, box_y_min - 20), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
                
                _, buffer = cv2.imencode('.jpg', img)
                image_base64 = base64.b64encode(buffer).decode('utf-8')
            except Exception as e:
                pass
        else:
            return jsonify({'error': 'Ảnh có quá nhiều cảnh vật dư thừa hoặc bị lóa. Vui lòng căn biển số lọt vào khung ngắm xanh lá!'}), 400

        # --- LOGIC RA/VÀO & TÍNH TIỀN ---
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status = "VÔ"
        time_in = current_time
        time_out = ""
        duration_str = ""
        fee = 0

        conn = sqlite3.connect(DB_PARKING)
        c = conn.cursor()
        
        c.execute("SELECT id, time_in FROM parking_logs WHERE plate=? AND status='VÔ' ORDER BY id DESC LIMIT 1", (plate_text,))
        row = c.fetchone()
        
        if row:
            log_id, db_time_in = row
            status = "RA"
            time_in = db_time_in
            time_out = current_time
            c.execute("UPDATE parking_logs SET time_out=?, status=? WHERE id=?", (time_out, status, log_id))
            
            try:
                t_in = datetime.strptime(time_in, "%Y-%m-%d %H:%M:%S")
                t_out = datetime.strptime(time_out, "%Y-%m-%d %H:%M:%S")
                delta = t_out - t_in
                minutes = int(delta.total_seconds() // 60)
                
                fee = minutes * 50
                duration_str = f"{minutes} phút"
                msg = f"Đậu {duration_str}. Thu {fee:,}đ."
            except Exception:
                msg = "Xe xuất bến thành công."
        else:
            c.execute("INSERT INTO parking_logs (plate, time_in, time_out, status) VALUES (?, ?, ?, ?)", 
                      (plate_text, time_in, "", status))
            log_id = c.lastrowid
            msg = "Đã lưu thông tin xe vào bãi."
            
        conn.commit()
        conn.close()

        with sheet_lock:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb.active
            if status == "VÔ":
                ws.append([log_id, plate_text, time_in, "", status])
            else:
                for r in range(2, ws.max_row + 1):
                    if ws.cell(row=r, column=1).value == log_id:
                        ws.cell(row=r, column=4, value=time_out)
                        ws.cell(row=r, column=5, value=status)
                        break
            wb.save(EXCEL_FILE)

        return jsonify({
            'plate': plate_text,
            'status': status,
            'time': current_time,
            'message': msg,
            'image_base64': image_base64,
            'duration': duration_str,
            'fee': fee
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ==========================================
# 3. LOGIC THỐNG KÊ DATA CHO WEB ADMIN
# ==========================================

def get_stats():
    try:
        conn = sqlite3.connect(DB_PARKING)
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM parking_logs WHERE status='VÔ'")
        inside = c.fetchone()[0]
        
        today_str = datetime.now().strftime("%Y-%m-%d")
        c.execute("SELECT COUNT(*) FROM parking_logs WHERE status='RA' AND time_out LIKE ?", (f"{today_str}%",))
        out_today = c.fetchone()[0]
        conn.close()
        return inside, out_today
    except:
        return 0, 0

def get_recent_logs(limit=50): 
    try:
        conn = sqlite3.connect(DB_PARKING)
        c = conn.cursor()
        c.execute("SELECT plate, time_in, time_out, status FROM parking_logs ORDER BY id DESC LIMIT ?", (limit,))
        rows = c.fetchall()
        conn.close()
        return rows
    except:
        return []

@app.route('/api/admin-data', methods=['GET'])
def get_admin_data():
    inside, out_today = get_stats()
    raw_logs = get_recent_logs()
    
    logs_in = []
    logs_out = []
    current_time = datetime.now()
    
    for log in raw_logs:
        plate, time_in, time_out, status = log
        if status == "VÔ":
            logs_in.append({
                'plate': plate,
                'time_in': time_in
            })
        elif status == "RA" and time_out:
            try:
                t_in = datetime.strptime(time_in, "%Y-%m-%d %H:%M:%S")
                t_out = datetime.strptime(time_out, "%Y-%m-%d %H:%M:%S")
                seconds_diff = (current_time - t_out).total_seconds()
                
                # Logic ẩn sau 30 giây giống hệt Tkinter cũ
                if seconds_diff <= 30:
                    minutes = int((t_out - t_in).total_seconds() // 60)
                    fee = minutes * 50
                    logs_out.append({
                        'plate': plate,
                        'duration': f"{minutes} phút",
                        'fee': f"{fee:,} VNĐ"
                    })
            except:
                pass
                
    return jsonify({
        'inside': inside,
        'out_today': out_today,
        'logs_in': logs_in,
        'logs_out': logs_out,
        'local_ip': get_local_ip()
    })

if __name__ == '__main__':
    # Chạy trực tiếp web server local
    app.run(host='0.0.0.0', port=5000, debug=True)
