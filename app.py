import sys
import os
import time
import sqlite3
import threading
import socket
import io
import re
import base64
import platform # MỚI: Thư viện lấy thông tin hệ điều hành/CPU
from datetime import datetime, timedelta
import requests

from flask import Flask, request, jsonify, render_template
from PIL import Image, ImageEnhance, ImageFilter
import openpyxl

import cv2
import numpy as np

# ==========================================
# 1. CẤU HÌNH HỆ THỐNG
# ==========================================

# MỚI: Ghi nhận thời điểm server bắt đầu khởi động để tính Uptime
START_TIME = time.time()

DATA_DIR = 'data'
TEMP_IMAGES_DIR = 'temp_images'
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(TEMP_IMAGES_DIR, exist_ok=True)

DB_PARKING = os.path.join(DATA_DIR, 'parking.db')

OCR_API_KEY = 'K84602471488957' 

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

init_db()
init_excel()

# ==========================================
# 2. FLASK ROUTES (ĐIỀU HƯỚNG WEB & API)
# ==========================================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/cam')
def camera_page():
    return render_template('cam.html')

@app.route('/admin')
def admin_page():
    return render_template('admin.html')

@app.route('/ping', methods=['POST'])
def ping():
    return jsonify({'status': 'ok'})

# ==========================================
# MỚI: ROUTE /device HIỂN THỊ INFO CPU ĐỂ GIỮ SERVER ONLINE
# ==========================================
@app.route('/device')
def device_info():
    uptime_seconds = int(time.time() - START_TIME)
    uptime_str = str(timedelta(seconds=uptime_seconds))
    
    system_info = {
        'os': platform.system(),
        'release': platform.release(),
        'cpu_cores': os.cpu_count(),
        'python_version': platform.python_version(),
        'server_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'uptime': uptime_str
    }
    
    # Giao diện HTML inline hiện đại (Dark Mode) không cần tạo thêm file template
    html_content = f"""
    <!DOCTYPE html>
    <html lang="vi">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Render Server Status</title>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-gray-900 text-white flex items-center justify-center min-h-screen font-sans">
        <div class="bg-gray-800 p-8 rounded-xl shadow-[0_0_20px_rgba(0,255,0,0.1)] w-full max-w-md border border-gray-700">
            <div class="flex items-center justify-between mb-6">
                <h1 class="text-2xl font-bold text-green-400">⚡ Server Status</h1>
                <span class="bg-green-500/20 text-green-400 text-xs px-3 py-1 rounded-full font-bold animate-pulse border border-green-500/30">ONLINE</span>
            </div>
            <div class="space-y-4">
                <div class="flex justify-between border-b border-gray-700 pb-3">
                    <span class="text-gray-400">Hệ điều hành:</span>
                    <span class="font-medium text-gray-200">{system_info['os']} {system_info['release']}</span>
                </div>
                <div class="flex justify-between border-b border-gray-700 pb-3">
                    <span class="text-gray-400">Cấu hình CPU:</span>
                    <span class="font-medium text-blue-300">{system_info['cpu_cores']} Cores</span>
                </div>
                <div class="flex justify-between border-b border-gray-700 pb-3">
                    <span class="text-gray-400">Môi trường:</span>
                    <span class="font-medium text-purple-400">Python {system_info['python_version']}</span>
                </div>
                <div class="flex justify-between border-b border-gray-700 pb-3">
                    <span class="text-gray-400">Giờ máy chủ:</span>
                    <span class="font-medium text-pink-300">{system_info['server_time']}</span>
                </div>
                <div class="flex justify-between border-b border-gray-700 pb-3">
                    <span class="text-gray-400">Thời gian chạy (Uptime):</span>
                    <span class="font-medium text-yellow-400">{system_info['uptime']}</span>
                </div>
            </div>
            <div class="mt-8 pt-4 border-t border-gray-700 text-center text-xs text-gray-500">
                Endpoint này dùng để theo dõi tài nguyên & giữ Render.com không bị sleep.
            </div>
        </div>
    </body>
    </html>
    """
    return html_content

@app.route('/upload-ocr', methods=['POST'])
def process_image():
    files = request.files.getlist('images')
    if not files:
        if 'image' in request.files:
            files = request.files.getlist('image')
        else:
            return jsonify({'error': 'Không tìm thấy ảnh'}), 400
    
    results = [] 

    for file in files:
        if file.filename == '':
            continue
            
        try:
            in_memory_file = io.BytesIO(file.read())
            file_bytes = np.frombuffer(in_memory_file.getvalue(), dtype=np.uint8)
            img = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)

            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
            sharpen_kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
            sharpened = cv2.filter2D(gray, -1, sharpen_kernel)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
            clahe_img = clahe.apply(sharpened)
            _, thresh = cv2.threshold(cv2.GaussianBlur(clahe_img, (5, 5), 0), 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3,3))
            morph = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
            kernel_small = cv2.getStructuringElement(cv2.MORPH_RECT, (2,2))
            open_morph = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel_small)

            plate_text = ""
            box_x_min, box_y_min, box_x_max, box_y_max = 9999, 9999, 0, 0
            image_base64 = ""

            encode_param = [int(cv2.IMWRITE_JPEG_QUALITY), 80]
            _, buffer = cv2.imencode('.jpg', img, encode_param)
            
            payload = {
                'apikey': OCR_API_KEY,
                'language': 'eng',            
                'isOverlayRequired': True,    
                'OCREngine': 2                
            }
            api_files = {'filename': ('plate.jpg', buffer.tobytes(), 'image/jpeg')}
            
            try:
                response = requests.post('https://api.ocr.space/parse/image', data=payload, files=api_files, timeout=15)
                result = response.json()
                
                if not result.get('IsErroredOnProcessing') and result.get('ParsedResults'):
                    parsed_result = result['ParsedResults'][0]
                    raw_text = parsed_result.get('ParsedText', '')
                    
                    clean_text = ''.join(e for e in raw_text if e.isalnum()).upper()
                    matches = re.findall(r'[1-9]\d[A-Z][A-Z0-9]?\d{4,5}', clean_text)
                    
                    if matches:
                        plate_text = max(matches, key=len)
                        
                        if parsed_result.get('TextOverlay'):
                            lines = parsed_result['TextOverlay'].get('Lines', [])
                            for line in lines:
                                line_text = ''.join(e for e in line.get('LineText', '') if e.isalnum()).upper()
                                if plate_text in line_text or line_text in plate_text:
                                    for word in line.get('Words', []):
                                        x = int(word['Left'])
                                        y = int(word['Top'])
                                        w = int(word['Width'])
                                        h = int(word['Height'])
                                        
                                        box_x_min = min(box_x_min, x)
                                        box_y_min = min(box_y_min, y)
                                        box_x_max = max(box_x_max, x + w)
                                        box_y_max = max(box_y_max, y + h)
            except Exception as api_err:
                results.append({'error': f'Lỗi kết nối API OCR đám mây: {str(api_err)}'})
                continue

            if plate_text:
                try:
                    if box_x_min != 9999:
                        cv2.rectangle(img, (box_x_min - 15, box_y_min - 15), (box_x_max + 15, box_y_max + 15), (0, 255, 0), 3)
                        cv2.putText(img, plate_text, (box_x_min - 15, box_y_min - 20), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
                    
                    _, buffer_final = cv2.imencode('.jpg', img)
                    image_base64 = base64.b64encode(buffer_final).decode('utf-8')
                except Exception as e:
                    pass
            else:
                results.append({'error': 'Ảnh bị lóa hoặc API không nhận diện được số. Vui lòng thử lại!'})
                continue

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

            results.append({
                'plate': plate_text,
                'status': status,
                'time': current_time,
                'message': msg,
                'image_base64': image_base64,
                'duration': duration_str,
                'fee': fee
            })

        except Exception as e:
            results.append({'error': str(e)})

    if len(files) == 1 and 'image' in request.files and 'images' not in request.files:
        if 'error' in results[0]:
            return jsonify(results[0]), 400
        return jsonify(results[0])
        
    return jsonify({'results': results})

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

# ==========================================
# 3. API ĐỒNG BỘ DỮ LIỆU REAL-TIME CHO WEB ADMIN
# ==========================================

@app.route('/api/admin-data', methods=['GET'])
def get_admin_data():
    inside, out_today = get_stats()
    logs = get_recent_logs()
    current_time = datetime.now()
    
    logs_in = []
    logs_out = []
    
    for log in logs:
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
                
                if seconds_diff <= 30:
                    minutes = int((t_out - t_in).total_seconds() // 60)
                    fee = minutes * 50
                    logs_out.append({
                        'plate': plate,
                        'duration': f"{minutes} phút",
                        'fee': f"{fee:,} VNĐ"
                    })
            except Exception:
                pass
                
    return jsonify({
        'inside': inside,
        'out_today': out_today,
        'logs_in': logs_in,
        'logs_out': logs_out,
        'local_ip': get_local_ip()
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False, use_reloader=False)
