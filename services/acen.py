# services/acen.py
from pathlib import Path
import pandas as pd

# Excel 상수 (constants 모듈 회피)
XL_CALC_MANUAL    = -4135
XL_CALC_AUTOMATIC = -4105

def extract_p_bi(file_storage):
    """
    업로드 엑셀 첫 번째 시트에서 P2~P10, BI2~BI10(총 9행) 추출.
    빈 문자열/공백은 None으로.
    """
    df = pd.read_excel(
        file_storage, sheet_name=0, usecols="P,BI",
        header=None, skiprows=1, nrows=9,
        engine="openpyxl", dtype=object
    )
    df.columns = ["P", "BI"]

    def norm(v):
        if pd.isna(v): return None
        if isinstance(v, str) and v.strip() == "": return None
        return v

    p_vals  = [norm(v) for v in df["P"].tolist()]
    bi_vals = [norm(v) for v in df["BI"].tolist()]
    return p_vals, bi_vals


def build_outputs_with_com(template_path: Path,
                           p_vals, bi_vals,
                           out_path: Path,
                           only_sheet_path: Path,
                           values_only=True):
    """
    Excel COM 한 번만 띄워서:
      1) 템플릿 열기(외부 링크 갱신 X, 계산/이벤트/화면 업데이트 OFF)
      2) Acen!B4/C4에 값 주입
      3) 전체 파일 SaveCopyAs(out_path)
      4) 'Acen 매출결의서' 시트만 새 통합문서로 복사 → SaveAs(only_sheet_path)
    """
    import pythoncom
    from win32com.client import DispatchEx

    pythoncom.CoInitialize()
    xl = DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try: xl.ScreenUpdating = False
    except: pass
    try: xl.EnableEvents = False
    except: pass
    try: xl.AskToUpdateLinks = False
    except: pass
    try: xl.Calculation = XL_CALC_MANUAL
    except: pass

    wb = None
    try:
        wb = xl.Workbooks.Open(str(template_path.resolve()), UpdateLinks=0)

        ws = next((s for s in wb.Worksheets if s.Name.strip().casefold() == "acen"), None)
        if ws is None:
            raise RuntimeError("시트 'Acen'을 찾을 수 없습니다.")

        start = 4
        n = max(len(p_vals), len(bi_vals))
        if n > 0:
            ws.Range(ws.Cells(start, 2), ws.Cells(start + n - 1, 3)).ClearContents()

        for i, v in enumerate(p_vals):
            ws.Cells(start + i, 2).Value2 = (None if v is None else v)
        for i, v in enumerate(bi_vals):
            ws.Cells(start + i, 3).Value2 = (None if v is None else v)

        wb.SaveAs(str(out_path.resolve()))

        ws2 = next((s for s in wb.Worksheets if s.Name.strip().casefold() == "acen 매출결의서"), None)
        if ws2 is None:
            raise RuntimeError("시트 'Acen 매출결의서'를 찾을 수 없습니다.")

        ws2.Copy()   # → ActiveWorkbook = 새 통합문서
        new_wb = xl.ActiveWorkbook
        new_ws = new_wb.Worksheets(1)

        if values_only:
            ur = new_ws.UsedRange
            ur.Value2 = ur.Value2
            # 외부 링크 끊기 (1 = xlLinkTypeExcelLinks)
            try:
                links = new_wb.LinkSources(1)
                if links:
                    for link in links:
                        new_wb.BreakLink(Name=link, Type=1)
            except Exception:
                pass
        # ★ 하드코딩: 저장 직전 시트 이름 고정
        new_ws.Name = "매출결의서"

        new_wb.SaveAs(str(only_sheet_path.resolve()), FileFormat=51)  # xlsx
        new_wb.Close(SaveChanges=False)
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try: xl.Calculation = XL_CALC_AUTOMATIC
        except: pass
        xl.Quit()
        pythoncom.CoUninitialize()
