# services/aicc.py
from pathlib import Path
from typing import List
import pandas as pd
from datetime import datetime
import pythoncom
from win32com.client import DispatchEx
import io, os, zipfile
import openpyxl

# Excel 계산 모드 상수
XL_CALC_MANUAL    = -4135
XL_CALC_AUTOMATIC = -4105

def extract_aicc_rows(file_storage) -> List[List[object]]:
    """
    업로드 엑셀 첫 시트의 A:M, 7행부터 데이터가 있는 부분까지 추출.
    완전 빈 행은 제거. 각 행 길이=13, 빈칸은 None.
    """
    df = pd.read_excel(
        file_storage, sheet_name=0, usecols="A:M",
        header=None, skiprows=6, engine="openpyxl", dtype=object
    ).dropna(how="all")

    def norm_cell(x):
        if pd.isna(x): return None
        if isinstance(x, str):
            s = x.strip()
            return None if s == "" else s
        return x

    df = df.applymap(norm_cell)
    return [list(row) for row in df.itertuples(index=False, name=None)]

def export_single_sheet(book_path: Path,
                        sheet_name: str,
                        dest_path: Path,
                        values_only: bool = True):
    """
    book_path에서 sheet_name 시트만 새 통합문서로 복사해 dest_path(.xlsx) 저장.
    values_only=True면 수식을 값으로 고정하고 외부링크 끊음.
    저장 직전 시트명은 하드코딩('매출결의서')으로 변경.
    """
 
    pythoncom.CoInitialize()
    xl = DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(book_path.resolve()), UpdateLinks=0, ReadOnly=True)

        ws = None
        for s in wb.Worksheets:
            if s.Name.strip().casefold() == sheet_name.strip().casefold():
                ws = s
                break
        if ws is None:
            raise RuntimeError(f"시트 '{sheet_name}'을(를) 찾을 수 없습니다.")

        ws.Copy()  # → ActiveWorkbook: 새 통합문서
        new_wb = xl.ActiveWorkbook
        new_ws = new_wb.Worksheets(1)

        if values_only:
            ur = new_ws.UsedRange
            ur.Value2 = ur.Value2
            try:
                links = new_wb.LinkSources(1)  # 1 = xlLinkTypeExcelLinks
                if links:
                    for link in links:
                        new_wb.BreakLink(Name=link, Type=1)
            except Exception:
                pass

        # 하드코딩 시트명
        new_ws.Name = "매출결의서"
        new_wb.SaveAs(str(dest_path.resolve()), FileFormat=51)
        new_wb.Close(SaveChanges=False)
    finally:
        try: wb.Close(SaveChanges=False)
        except: pass
        try: xl.Quit()
        except: pass
        pythoncom.CoUninitialize()

def inject_and_calculate_with_com(book_path: Path, rows: list) -> int:
    """
    [win32com] 데이터를 주입하고, 통합 문서를 재계산한 뒤,
    모든 변경사항을 파일에 저장하고 종료합니다.
    """
    dbg = lambda msg: print(f"[AICC][{datetime.now():%H:%M:%S}] {msg}", flush=True)
    pythoncom.CoInitialize()
    xl = DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    
    wb = None
    written_count = 0
    try:
        wb = xl.Workbooks.Open(str(book_path.resolve()))
        
        # 1. 데이터 주입
        xl.Calculation = -4135 # 수동 계산
        ws_aicc = wb.Worksheets("AICC")
        start_row, cols = 3, 13
        data2d = tuple(tuple(row[i] if i < len(row) else None for i in range(cols)) for row in rows)
        ws_aicc.Range(ws_aicc.Cells(start_row, 1), ws_aicc.Cells(start_row + len(rows) - 1, cols)).Value2 = data2d
        written_count = len(rows)
        dbg(f"INFO: {written_count}행 데이터 주입 완료.")

        # 2. 재계산
        xl.Calculation = -4105 # 자동 계산
        xl.Calculate()
        dbg("INFO: 통합 문서 재계산 완료.")
        
    finally:
        if wb:
            # ✨ 핵심 변경점: 모든 작업을 마친 후 저장하고 닫습니다.
            wb.Close(SaveChanges=True)
            dbg("INFO: win32com 작업 완료 및 파일 저장.")
        if xl:
            xl.Quit()
        pythoncom.CoUninitialize()
        
    return written_count

def read_and_inject_with_openpyxl(book_path: Path, now: datetime) -> Path:
    """
    [openpyxl] 계산된 파일에서 값을 읽어 최종 주입하고,
    '다른 이름으로 저장'하여 파일 잠금 문제를 회피합니다.
    - 반환: 최종 저장된 파일의 경로(Path 객체)
    """
    dbg = lambda msg: print(f"[AICC][{datetime.now():%H:%M:%S}] {msg}", flush=True)
    
    wb = openpyxl.load_workbook(book_path, data_only=True)
    ws_main = wb.worksheets[0]
    ws_calc = wb["매출계산기"]

    # --- ✨ 키 값 비교 디버깅 로직 ---
    # 1. '매출계산기' C열(Key) 집합 생성
    calc_keys = set()
    for row in ws_calc.iter_rows(min_row=3, max_row=50, min_col=3, max_col=3, values_only=True):
        key = row[0]
        if key is None: break
        calc_keys.add(str(key).strip())
    
    # 2. '2025' 시트 B열(Key) 집합 생성
    main_keys = set()
    for cell in ws_main['B'][4:24]: # B5:B24
        if cell.value is None: continue
        main_keys.add(str(cell.value).strip())

    dbg("--- [키 값 비교 시작] ---")
    dbg(f"매출계산기 키 개수: {len(calc_keys)}")
    dbg(f"'2025' 시트 키 개수: {len(main_keys)}")
    unmatched_in_main = main_keys - calc_keys
    if unmatched_in_main:
        dbg(f"WARN: '2025' 시트에는 있지만 '매출계산기'에 없는 키: {unmatched_in_main}")
    dbg("--- [키 값 비교 종료] ---")
    # --- ✨ 디버깅 종료 ---

    # 기존 로직: 딕셔너리 생성 및 값 주입
    results_map = {str(k).strip(): v for k, v in ws_calc.iter_rows(min_row=3, max_row=50, min_col=3, max_col=4, values_only=True) if k is not None}

    current_month_str = f"{now.month}월"
    target_col_number = None
    for col_idx, cell in enumerate(ws_main[4][11:23], start=12):
        if str(cell.value).strip() == current_month_str:
            target_col_number = col_idx
            break
            
    if target_col_number:
        dbg(f"INFO: openpyxl로 '{current_month_str}'에 해당하는 {target_col_number}열에 데이터 주입 시작...")
        match_count = 0
        for row_idx in range(5, 25):
            lookup_key = str(ws_main.cell(row=row_idx, column=2).value).strip()
            value_to_write = results_map.get(lookup_key)
            if value_to_write is not None:
                ws_main.cell(row=row_idx, column=target_col_number, value=value_to_write)
                match_count += 1
        dbg(f"INFO: openpyxl로 데이터 주입 완료 (총 {match_count}개 매칭).")
    
    # ✨ --- 핵심 변경점: 다른 이름으로 저장 --- ✨
    wb.save(book_path)
    dbg(f"INFO: openpyxl 작업 완료 및 새 파일 '{book_path.name}'으로 저장.")
    
    return book_path # 새로 저장된 파일의 경로를 반환

def export_and_zip(book_path: Path, ymd_dot: str) -> io.BytesIO:
    """최종 파일을 읽어 시트를 추출하고 ZIP 파일의 BytesIO 객체를 반환합니다."""
    # export_single_sheet 함수는 별도로 구현되어 있어야 합니다.
    only_sheet_path = book_path.parent / f"매출결의서_KT AICC_{ymd_dot}.xlsx"
    export_single_sheet(book_path, "AICC 매출결의서", only_sheet_path, values_only=True)

    month_bytes = book_path.read_bytes()
    only_bytes  = only_sheet_path.read_bytes()
    
    # 임시 단일 시트 파일 삭제
    try:
        if only_sheet_path.exists():
            os.remove(only_sheet_path)
    except Exception:
        pass

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(book_path.name, month_bytes)
        zf.writestr(only_sheet_path.name, only_bytes)
    buf.seek(0)
    return buf