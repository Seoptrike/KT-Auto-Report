from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from pathlib import Path
from datetime import datetime, timedelta
import io, os, zipfile
import time

from services.acen import extract_p_bi, build_outputs_with_com
from services.aicc import (
    extract_aicc_rows,
    export_single_sheet,
    inject_and_calculate_with_com,
    read_and_inject_with_openpyxl
)

app = Flask(__name__)
app.config["SECRET_KEY"] = "change-me-to-a-real-secret-key"

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "업무실적계산기.xlsx"  # 원본 템플릿 파일 이름


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


# -------------------- ACEN --------------------
@app.post("/run/acen")
def run_acen():
    acen_file = request.files.get("acen_file")
    if not acen_file or not acen_file.filename:
        flash("ACEN 파일을 선택하세요.", "error")
        return redirect(url_for("index"))

    try:
        now = datetime.now()
        paths = get_current_and_previous_month_paths(now)
        source_path = get_source_path(paths["prev_month_book"], TEMPLATE_PATH)

        if not source_path.exists():
            raise FileNotFoundError(f"읽을 원본 파일이 없습니다: {source_path}")

        p_vals, bi_vals = extract_p_bi(acen_file)

        out_path = paths["current_month_book"]
        only_sheet_path = out_path.parent / f"Acen_매출결의서_{paths['stamp']}.xlsx"

        build_outputs_with_com(
            source_path, p_vals, bi_vals, out_path, only_sheet_path, values_only=True
        )

        full_bytes = out_path.read_bytes()
        sheet_bytes = only_sheet_path.read_bytes()
        os.remove(only_sheet_path)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(out_path.name, full_bytes)
            zf.writestr(
                f"매출결의서_KT Acen_유지수수료_{paths['ymd_dot']}.xlsx", sheet_bytes
            )
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=f"업무실적_{paths['stamp']}.zip",
            mimetype="application/zip",
            max_age=0
        )
    except Exception as e:
        flash(f"ACEN 처리 오류: {e}", "error")
        return redirect(url_for("index"))


# -------------------- AICC --------------------
@app.post("/run/aicc")
def run_aicc():
    files = [f for f in request.files.getlist("aicc_files") if f and f.filename]
    if not files:
        flash("AICC 파일들을 선택하세요.", "error")
        return redirect(url_for("index"))

    try:
        now = datetime.now()
        paths = get_current_and_previous_month_paths(now)
        month_book = paths["current_month_book"]

        if not month_book.exists():
            raise FileNotFoundError(
                f"서버에 {month_book.name}이 없습니다. 먼저 ACEN 실행으로 생성하세요."
            )

        rows = []
        for fs in files:
            try:
                fs.stream.seek(0)
            except:
                pass
            rows.extend(extract_aicc_rows(fs))

        if not rows:
            flash("AICC에서 추출된 데이터가 없습니다.", "warning")
            return redirect(url_for("index"))

        # 1. COM으로 데이터 주입 및 계산
        written = inject_and_calculate_with_com(month_book, rows)
        if written == 0:
            raise RuntimeError("주입된 데이터가 없습니다.")

        time.sleep(2)  # 파일 I/O 안정성을 위한 대기

        # 2. openpyxl로 최종 결과 주입
        read_and_inject_with_openpyxl(month_book, now)

        # 3. 시트 추출 및 ZIP 생성
        only_sheet_path = month_book.parent / f"매출결의서_KT AICC_{paths['ymd_dot']}.xlsx"
        export_single_sheet(month_book, "AICC 매출결의서", only_sheet_path, values_only=True)

        month_bytes = month_book.read_bytes()
        only_bytes = only_sheet_path.read_bytes()
        os.remove(only_sheet_path)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(month_book.name, month_bytes)
            zf.writestr(only_sheet_path.name, only_bytes)
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=f"AICC_{paths['stamp']}.zip",
            mimetype="application/zip",
            max_age=0
        )
    except Exception as e:
        import traceback
        print(traceback.format_exc(), flush=True)
        flash(f"AICC 처리 오류: {e}", "error")
        return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)
