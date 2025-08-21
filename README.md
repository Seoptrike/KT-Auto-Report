# 📊 매출결의서 작성 자동화 프로그램

월별로 반복되는 ACEN 및 AICC 정산 자료를 자동으로 취합하여, 업무실적 보고서를 생성하는 데스크톱 애플리케이션입니다.

---

## 🚀 다운로드

👉 [최신 실행 파일 다운로드](https://github.com/Seoptrike/KT-Auto-Report/releases/latest)

[1.0.0 다운로드] (https://github.com/Seoptrike/KT-Auto-Report/releases/download/v1.0.0/KT_auto_report_v1.0.exe)

---

매월 반복되는 정산 업무는 수작업으로 진행되어 시간이 오래 걸리고 실수가 잦았습니다. 이 문제를 해결하기 위해 자동화 프로그램을 개발했습니다.

처음에는 Flask 기반 웹 서비스를 목표로 했으나, 개발 과정에서 마주친 기술적 문제들을 해결하며 사용자 PC에서 직접 실행하는 데스크톱 애플리케이션으로 최종 완성되었습니다.

주요 개발 과정은 다음과 같습니다.

문제 1: Excel 수식 손상

원인: openpyxl 라이브러리가 기존 파일의 복잡한 수식을 유지하지 못했습니다.

해결: MS Office를 직접 제어하는 win32com을 도입하여 수식과 서식을 온전히 보존했습니다.

문제 2: 서버 배포의 한계

원인: win32com은 MS Office가 설치된 환경에서만 동작하므로, Office가 없는 사내 서버(VM) 배포가 불가능했습니다.

해결: PyInstaller와 pywebview를 사용해, 누구나 자신의 PC에서 쉽게 실행할 수 있는 .exe 데스크톱 앱으로 방향을 전환했습니다.

문제 3: UI 멈춤 현상

원인: 무거운 Excel 작업이 실행되는 동안 프로그램 전체가 멈춰 사용자가 진행 상황을 알 수 없었습니다.

해결: threading과 queue를 적용하여 백그라운드 작업과 UI를 분리하고, 처리 과정을 실시간 로그로 보여주어 사용자 경험을 개선했습니다.

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
