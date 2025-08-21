from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from pathlib import Path
from datetime import datetime, timedelta
import pythoncom
import io, os, zipfile
import time
import webview 
import sys
import queue  # 👈 1. queue 임포트
import threading # 👈 2. threading 임포트

# --- 전역 변수로 Queue(우체통) 생성 ---
log_queue = queue.Queue()

# --- 3. 로그를 화면에 업데이트하는 별도의 함수 ---
def log_updater(window):
    """Queue에 메시지가 들어오면 GUI에 안전하게 업데이트합니다."""
    while True:
        message = log_queue.get() # 메시지가 들어올 때까지 여기서 기다림
        if message is None: # 종료 신호
            break
        
        # JavaScript 코드를 실행해서 log-box의 내용을 업데이트
        escaped_message = message.replace('\\', '\\\\').replace('"', '\\"').replace("'", "\\'").replace('\n', '\\n')
        js_code = f"""
            var logBox = document.getElementById('log-box');
            logBox.innerHTML += '{escaped_message}\\n';
            logBox.scrollTop = logBox.scrollHeight;
        """
        window.evaluate_js(js_code)
        time.sleep(0.01)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


from services.acen import extract_p_bi, build_outputs_with_com
from services.aicc import extract_aicc_rows, export_single_sheet, inject_and_calculate_with_com, read_and_inject_with_openpyxl

# 👈 1. BASE_DIR 정의 변경: .exe 실행 위치를 기준으로 삼음
if getattr(sys, 'frozen', False):
    # PyInstaller로 패키징된 경우 (.exe)
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    # 일반 파이썬으로 실행된 경우 (.py)
    BASE_DIR = Path(__file__).resolve().parent

# 👈 2. 템플릿 파일 경로는 resource_path 사용
TEMPLATE_PATH = Path(resource_path("업무실적계산기.xlsx"))

# 👈 3. Flask 앱 생성 시 templates 폴더 경로 명시
app = Flask(__name__, template_folder=resource_path('templates'))
app.config["SECRET_KEY"] = "change-me-to-a-real-secret-key"

# BASE_DIR = Path(__file__).resolve().parent
# TEMPLATE_PATH = BASE_DIR / "업무실적계산기.xlsx" # 원본 템플릿 파일 이름

# --- 공통 Helper 함수 ---

def get_current_and_previous_month_paths(now: datetime):
    """현재 달과 이전 달의 경로 정보를 생성하여 반환합니다."""
    year_dir = BASE_DIR / now.strftime("%Y")
    ym_dir = year_dir / now.strftime("%y.%m")
    ym_dir.mkdir(parents=True, exist_ok=True)
    stamp = now.strftime("%Y%m")
    current_month_book = ym_dir / f"업무실적_{stamp}.xlsx"
    
    first_day_of_current_month = now.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    prev_month_year = last_day_of_previous_month.strftime("%Y")
    prev_month_ym = last_day_of_previous_month.strftime("%y.%m")
    prev_month_stamp = last_day_of_previous_month.strftime("%Y%m")
    prev_month_dir = BASE_DIR / prev_month_year / prev_month_ym
    prev_month_book = prev_month_dir / f"업무실적_{prev_month_stamp}.xlsx"
    
    return {
        "ymd_dot": now.strftime("%y.%m.%d"),
        "stamp": stamp,
        "current_month_book": current_month_book,
        "prev_month_book": prev_month_book
    }

def get_source_path(prev_month_book: Path, template_path: Path) -> Path:
    """이전 달 파일이 있으면 사용하고, 없으면 기본 템플릿을 반환합니다."""
    if prev_month_book.exists():
        print(f"INFO: 이전 달 파일 '{prev_month_book.name}'을 원본으로 사용합니다.")
        return prev_month_book
    else:
        print(f"INFO: 이전 달 파일이 없어 기본 템플릿 '{template_path.name}'을 사용합니다.")
        return template_path
    

# --- Flask 라우트 ---

@app.get("/")
def index():
    return render_template("index.html")

# --- ACEN 작업자 함수 (새로 만들기) ---
def acen_worker(file_content):
    """실제 엑셀 처리를 담당하는 함수 (별도 스레드에서 실행됨)"""
    pythoncom.CoInitialize()
    try:
        log_queue.put("ACEN 처리 시작...")
        now = datetime.now()
        paths = get_current_and_previous_month_paths(now)
        log_queue.put("이전 달 데이터 확인 중...")
        source_path = get_source_path(paths["prev_month_book"], TEMPLATE_PATH)

        if not source_path.exists():
            raise FileNotFoundError(f"읽을 원본 파일이 없습니다: {source_path}")
        
        log_queue.put("엑셀 데이터 추출 중...")
        # 파일 내용을 메모리에서 직접 읽도록 수정
        p_vals, bi_vals = extract_p_bi(io.BytesIO(file_content))
        out_path = paths["current_month_book"]
        only_sheet_path = out_path.parent / f"Acen_매출결의서_{paths['stamp']}.xlsx"
        
        log_queue.put("엑셀 파일 생성 및 계산 실행 중 (시간이 걸릴 수 있습니다)...")
        build_outputs_with_com(source_path, p_vals, bi_vals, out_path, only_sheet_path, values_only=True)
        
        log_queue.put("결과 파일 압축 중...")
        full_bytes = out_path.read_bytes()
        sheet_bytes = only_sheet_path.read_bytes()
        os.remove(only_sheet_path)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(out_path.name, full_bytes)
            zf.writestr(f"매출결의서_KT Acen_유지수수료_{paths['ymd_dot']}.xlsx", sheet_bytes)
        buf.seek(0)

        log_queue.put("다운로드 폴더에 파일 저장 중...")
        downloads_path = Path.home() / "Downloads"
        downloads_path.mkdir(exist_ok=True)
        file_name = f"업무실적_{paths['stamp']}.zip"
        save_path = downloads_path / file_name
        with open(save_path, "wb") as f:
            f.write(buf.getbuffer())

        log_queue.put(f"✅ 작업 완료! '다운로드' 폴더에 '{file_name}'이 저장되었습니다.")
        # 작업 완료 후에는 flash 메시지를 직접 보낼 수 없으므로 로그로 대체
    except Exception as e:
        log_queue.put(f"❌ ACEN 처리 오류 발생: {e}")
    finally:
        pythoncom.CoUninitialize()

# -------------------- ACEN --------------------
@app.post("/run/acen")
def run_acen():
    acen_file = request.files.get("acen_file")
    if not acen_file or not acen_file.filename:
        flash("ACEN 파일을 선택하세요.", "error")
        return redirect(url_for("index"))

    # 파일 내용을 읽어서 worker 함수에 전달
    file_content = acen_file.read()
    
    # 별도의 스레드를 생성하여 worker 함수를 실행
    worker_thread = threading.Thread(target=acen_worker, args=(file_content,))
    worker_thread.start()
    
    # "작업 시작됨"을 알리고 즉시 페이지 새로고침
    flash("ACEN 작업이 백그라운드에서 시작되었습니다. 잠시 후 로그를 확인하세요.")
    return redirect(url_for("index"))

# --- AICC 작업자 함수 (새로 만들기) ---
def aicc_worker(files_content):
    """실제 AICC 엑셀 처리를 담당하는 함수 (별도 스레드에서 실행됨)"""
    pythoncom.CoInitialize()
    try:
        log_queue.put("AICC 처리 시작...")
        now = datetime.now()
        paths = get_current_and_previous_month_paths(now)
        month_book = paths["current_month_book"]

        if not month_book.exists():
            log_queue.put(f"❌ 오류: {month_book.name} 파일을 찾을 수 없습니다. ACEN을 먼저 실행하세요.")
            raise FileNotFoundError(f"서버에 {month_book.name}이 없습니다. 먼저 ACEN 실행으로 생성하세요.")

        log_queue.put(f"{len(files_content)}개 AICC 파일에서 데이터 추출 중...")
        rows = []
        for file_data in files_content:
            rows.extend(extract_aicc_rows(io.BytesIO(file_data)))
        
        if not rows:
            log_queue.put("⚠️ 경고: AICC 파일에서 추출된 데이터가 없습니다.")
            # flash는 여기서 직접 사용할 수 없으므로 로그로만 남깁니다.
            return # 작업 종료

        log_queue.put("데이터 주입 및 계산 실행 중 (COM)...")
        written = inject_and_calculate_with_com(month_book, rows)
        if written == 0:
            log_queue.put("❌ 오류: COM을 통해 데이터가 주입되지 않았습니다.")
            raise RuntimeError("주입된 데이터가 없습니다.")
        
        time.sleep(2) # 파일 I/O 안정성을 위한 대기

        log_queue.put("최종 결과 반영 중 (openpyxl)...")
        read_and_inject_with_openpyxl(month_book, now)
        
        log_queue.put("결과 시트 추출 중...")
        only_sheet_path = month_book.parent / f"매출결의서_KT AICC_{paths['ymd_dot']}.xlsx"
        export_single_sheet(month_book, "AICC 매출결의서", only_sheet_path, values_only=True)
        
        log_queue.put("결과 파일 압축 중...")
        month_bytes = month_book.read_bytes()
        only_bytes  = only_sheet_path.read_bytes()
        os.remove(only_sheet_path)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(month_book.name, month_bytes)
            zf.writestr(only_sheet_path.name, only_bytes)
        buf.seek(0)
        
        log_queue.put("다운로드 폴더에 파일 저장 중...")
        downloads_path = Path.home() / "Downloads"
        downloads_path.mkdir(exist_ok=True)
        
        file_name = f"AICC+Acen업무실적_{paths['stamp']}.zip"
        save_path = downloads_path / file_name

        with open(save_path, "wb") as f:
            f.write(buf.getbuffer())

        log_queue.put(f"✅ 작업 완료! '다운로드' 폴더를 확인하세요.")
        # flash는 다른 스레드에서 직접 호출할 수 없습니다.

    except Exception as e:
        import traceback
        log_queue.put(f"❌ AICC 처리 중 심각한 오류 발생: {e}")
        print(traceback.format_exc(), flush=True)
    finally:
        pythoncom.CoUninitialize()

# --- AICC 라우트 함수 (수정) ---
@app.post("/run/aicc")
def run_aicc():
    files = [f for f in request.files.getlist("aicc_files") if f and f.filename]
    if not files:
        flash("AICC 파일들을 선택하세요.", "error")
        return redirect(url_for("index"))

    # 파일 내용을 미리 읽어 리스트에 저장
    files_content = [f.read() for f in files]
    
    # 별도의 스레드를 생성하여 worker 함수를 실행
    worker_thread = threading.Thread(target=aicc_worker, args=(files_content,))
    worker_thread.start()
    
    # "작업 시작됨"을 알리고 즉시 페이지 새로고침
    flash("AICC 작업이 백그라운드에서 시작되었습니다. 잠시 후 로그를 확인하세요.")
    return redirect(url_for("index"))

if __name__ == "__main__":
    print("프로그램 시작 중...")
    
    window = webview.create_window("업무 실적 계산기", app)
    
    # 👈 5. log_updater 함수를 별도 스레드로 실행
    t = threading.Thread(target=log_updater, args=(window,))
    t.daemon = True
    t.start()
    
    webview.start()