# app.py (lean)
from __future__ import annotations
import os
from pathlib import Path
from datetime import datetime
from io import BytesIO
import zipfile

from flask import (
    Flask, render_template, request, redirect,
    url_for, flash, send_file
)

# --- 프로젝트 루트 import 경로 ---
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# --- 서비스 로직 ---
from services.acen import run_acen_pipeline
from services.aicc import (
    combine_bghm_from_paths,
    enrich_bghm_rows,
    group_sum_by_name_title_H,
    map_grouped_names,
    write_to_excel,
)
from services.sum import (
    find_latest_file_for_month,
    extract_D_K_rows,
    extract_company_aicc,
    extract_company_acen,
    merge_by_company,
    apply_sum_name_mapping,
    fill_sum_template,
)

# -----------------------------
# Flask 기본 설정
# -----------------------------
app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB 업로드 제한(필요시 조절)

# 디렉토리 상수
BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"

# 템플릿 경로
AICC_TEMPLATE = TEMPLATES_DIR / "AICC 매출결의서.xlsx"
ACEN_TEMPLATE = TEMPLATES_DIR / "Acen 매출결의서.xlsx"
SUM_TEMPLATE  = TEMPLATES_DIR / "업무실적.xlsx"

# 허용 확장자
ALLOWED_EXTS = {".xlsx", ".xlsm"}
def _is_allowed(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTS

# -----------------------------
# Routes
# -----------------------------
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

# ACEN: 파일 업로드 → 바로 XLSX 응답
@app.route("/run/acen", methods=["POST"])
def run_acen():
    f = request.files.get("acen_file")
    if not f or f.filename == "" or not _is_allowed(f.filename):
        flash("유효한 ACEN 파일을 선택해주세요.", "error")
        return redirect(url_for("index"))
    
    # ★ report_day 입력값 읽기
    rd_str = request.form.get("report_day", "").strip()
    report_day = None
    if rd_str.isdigit():
        report_day = int(rd_str)
        
    data = f.read()
    try:
        out_path = run_acen_pipeline(
            file_like=BytesIO(data),
            template_path=ACEN_TEMPLATE,
            base_dir=OUTPUT_DIR,
            write_formulas=False,
            date_fmt="dots",
            report_day=report_day,
        )
        return send_file(
            Path(out_path),
            as_attachment=True,
            download_name=Path(out_path).name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        app.logger.exception(e)
        flash(f"처리 중 오류 발생: {e}", "error")
        return redirect(url_for("index"))

# AICC: 파일 업로드 → AICC XLSX + 업무실적 XLSX를 ZIP으로 묶어 바로 응답
@app.route("/run/aicc", methods=["POST"])
def run_aicc():
    files = request.files.getlist("aicc_files")
    if not files:
        flash("AICC 파일을 선택해주세요.", "error")
        return redirect(url_for("index"))

    streams = []
    for f in files:
        if f and f.filename and _is_allowed(f.filename):
            streams.append(f.stream)
        else:
            flash(f"허용되지 않는 파일 형식: {getattr(f, 'filename', '')}", "error")
            return redirect(url_for("index"))
        
    rd_str = request.form.get("report_day", "").strip()
    report_day = None
    if rd_str.isdigit():
        report_day = int(rd_str)

    try:
        # 1) AICC 매출결의서 생성
        rows, settlement_month = combine_bghm_from_paths(streams, start_row=7)
        enriched = enrich_bghm_rows(rows)
        grouped  = group_sum_by_name_title_H(enriched, keep_order=True)
        mapped   = map_grouped_names(grouped)
        aicc_out = write_to_excel(
            mapped,
            AICC_TEMPLATE,
            base_dir=OUTPUT_DIR,
            settlement_month=settlement_month,
            report_day=report_day,
        )

        # === 전달(=정산월) 기준으로 파일 선정 ===
        month_basis = settlement_month or datetime.now()

        # 2) 정산월 기준 AICC/ACEN 파일 집계 → 업무실적 업데이트
        #    AICC는 방금 만든 aicc_out을 그대로 사용 (같은 정산월 폴더에 저장됨)
        latest_aicc = Path(aicc_out)

        #    ACEN은 정산월(YYYY/MM) 폴더에서 prefix로 검색
        latest_acen = find_latest_file_for_month(
            OUTPUT_DIR, month_basis, prefix="매출결의서_KT ACen"
        )

        merged_input = []
        if latest_aicc and latest_aicc.exists():
            for d, k in extract_D_K_rows(latest_aicc, start_row=11, end_row=25):
                name = extract_company_aicc(d)
                if name:
                    merged_input.append((name, k))

        if latest_acen and latest_acen.exists():
            for d, k in extract_D_K_rows(latest_acen, start_row=11, end_row=25):
                name = extract_company_acen(d)
                if name:
                    merged_input.append((name, k))

        merged     = merge_by_company(merged_input)
        mapped_sum = apply_sum_name_mapping(merged)
        sum_out = fill_sum_template(
            mapped_sum,
            SUM_TEMPLATE,
            out_base_dir=OUTPUT_DIR,
            settlement_month=settlement_month,  # ★ 전달
            report_day=report_day,              # ★ 전달
        )

        # 3) ZIP으로 묶어 즉시 다운로드 (파일명도 정산월 기준으로)
        zip_name = f"KT업무실적_{month_basis:%Y.%m}.zip"
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.write(Path(aicc_out), arcname=Path(aicc_out).name)
            if latest_acen and latest_acen.exists():
                zf.write(Path(sum_out),  arcname=Path(sum_out).name)
            else:
                zf.write(Path(sum_out),  arcname=Path(sum_out).name)  # ACEN 없어도 sum_out은 포함
        buf.seek(0)

        return send_file(buf, as_attachment=True, download_name=zip_name, mimetype="application/zip")

    except Exception as e:
        app.logger.exception(e)
        flash(f"처리 중 오류 발생: {e}", "error")
        return redirect(url_for("index"))

if __name__ == "__main__":
    # 개발용 실행 (프로덕션은 gunicorn/uwsgi 권장)
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
