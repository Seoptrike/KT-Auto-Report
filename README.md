# 📊 매출결의서 작성 자동화 프로그램

월별로 반복되는 ACEN 및 AICC 정산 자료를 자동으로 취합하여, 업무실적 보고서를 생성하는 데스크톱 애플리케이션입니다.

---

## 🚀 다운로드

👉 [최신 실행 파일 다운로드](https://github.com/Seoptrike/KT-Auto-Report/releases/latest)

[1.0.0 다운로드] (https://github.com/Seoptrike/KT-Auto-Report/releases/download/v1.0.0/KT_auto_report_v1.0.exe)

---

## 📌 프로젝트 개요

매월 반복되는 정산 업무를 수작업으로 진행하다 보니 **시간이 오래 걸리고
실수가 잦은 문제**가 있었습니다.\
이를 해결하기 위해, **정산 자동화 프로그램**을 직접 개발했습니다.

처음에는 **Flask 기반 웹 서비스**로 시작했지만,\
여러 기술적 제약을 해결하는 과정에서\
최종적으로 **사용자 PC에서 실행 가능한 데스크톱 애플리케이션(.exe)**
으로 완성되었습니다.

---

## 📸 스크린샷

![프로그램 실행 화면](images/main1.png)

---

## 🚀 주요 기능

- 💻 데스크톱 GUI: 웹 기술(HTML, CSS, Bootstrap) 기반  
- 📂 엑셀 자동 처리: ACEN 및 다중 AICC 파일 업로드 → 자동 통합  
- ⚡ 자동 연산: win32com을 활용한 업무실적 파일 업데이트 및 계산  
- 📜 작업 로그 표시: 비동기 처리 기반의 실시간 로그 확인  
- 📦 결과물 자동 저장: 최종 결과물(.zip)을 다운로드 폴더에 저장  
- 🔨 손쉬운 배포: 단일 실행 파일(.exe) 빌드 및 배포 지원  

---

## 🛠️ 기술 스택

- 언어: Python  
- 백엔드: Flask  
- 프론트엔드: HTML, CSS (Bootstrap 5)  
- 데스크톱 앱: pywebview  
- 엑셀 자동화: win32com, openpyxl  
- 동시성 처리: threading, queue  
- 배포: PyInstaller  
---

## ⚡ 개발 과정에서의 주요 이슈와 해결 방법

### 🧩 문제 1: Excel 수식 손상

-   **원인**: `openpyxl` 라이브러리가 복잡한 Excel 수식을 보존하지 못함\
-   **해결**: `win32com` 으로 MS Office를 직접 제어하여\
    → **수식과 서식을 온전히 보존**


### 🖥️ 문제 2: 서버 배포 한계

-   **원인**: `win32com`은 MS Office가 설치된 환경에서만 동작 → Office가
    없는 사내 VM에 배포 불가\
-   **해결**: `PyInstaller` + `pywebview` 활용\
    → **누구나 실행 가능한 .exe 데스크톱 앱**으로 전환



### 🚦 문제 3: UI 멈춤 현상

-   **원인**: 무거운 Excel 작업 실행 시 프로그램 전체가 멈춰, 진행 상황
    확인 불가\
-   **해결**: `threading` + `queue` 적용\
    → **백그라운드 처리 + 실시간 로그 출력**으로 사용자 경험 개선



## ✅ 결과

-   반복적인 정산 업무 시간을 **대폭 단축**\
-   **사람이 하던 단순 반복 실수 제거**\
-   IT 환경 제약 없이 **사내 누구나 쉽게 실행 가능**

---

## ⚙️ 로컬 실행 방법

1. 저장소 클론
    git clone https://github.com/Seoptrike/KT-Auto-Report.git
    cd KT-Auto-Report

2. 가상환경 생성 및 활성화 (권장)
    python -m venv venv
    source venv/bin/activate      # macOS/Linux
    .\venv\Scripts\activate       # Windows

3. 의존성 설치
    pip install -r requirements.txt

4. 개발 모드 실행
    python develop.py

---

## 📦 EXE 파일로 빌드하기

PyInstaller를 이용하여 단일 실행 파일(.exe)을 생성할 수 있습니다.

    pyinstaller --onefile --windowed \
        --hidden-import=win32com.client \
        --add-data "templates;templates" \
        --add-data "업무실적계산기.xlsx;." \
        --add-data "services;services" \
        app.py

- 생성된 실행 파일은 dist/ 디렉토리에서 확인할 수 있습니다.

---

## 📝 라이선스

이 프로젝트는 MIT License를 따릅니다.
