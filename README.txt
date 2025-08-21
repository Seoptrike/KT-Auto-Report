# 📊 매출결의서 작성 자동화 프로그램

월별로 반복되는 ACEN, AICC 정산 자료 취합 및 업무실적 파일 생성을 자동화하여 업무 효율을 극대화하는 데스크톱 애플리케이션입니다.

## ✨ 프로젝트 소개

매월 반복되는 정산 업무는 여러 개의 Excel 파일을 수작업으로 취합해야 하므로 시간이 많이 소요되고, 잦은 인적 오류(Human Error)가 발생하는 문제점이 있었습니다. 이 프로젝트는 해당 과정을 자동화하여 업무 시간을 분 단위로 단축하고 데이터의 정확성을 100% 보장하기 위해 개발되었습니다.

## 📸 스크린샷
![프로그램 실행 화면](images/main1.png)

*(이곳에 프로그램 실행 화면을 캡처한 이미지 파일(예: `screenshot.png`)을 넣으세요. 깃허브에 이미지 파일을 함께 올린 후 `![스크린샷](screenshot.png)` 와 같은 형식으로 링크하면 됩니다.)*

## 🚀 주요 기능

- 웹 기술(HTML, CSS) 기반의 데스크톱 GUI 제공
- ACEN 및 다중 AICC 엑셀 파일 업로드 및 데이터 자동 통합
- `win32com`을 이용한 업무실적 파일 자동 업데이트 및 계산
- 비동기 처리를 통한 실시간 작업 로그 표시 기능
- 최종 결과물(zip)을 '다운로드' 폴더에 자동 저장
- 단일 실행 파일(`.exe`) 생성 및 배포 지원

## 🛠️ 사용 기술

- 언어: Python
- 주요 라이브러리: Flask, pywebview, win32com, openpyxl
- 패키징: PyInstaller
- 프론트엔드: HTML, CSS (Bootstrap 5)

## ⚙️ 로컬에서 실행하기

1.  저장소 복제
    ```bash
    git clone [https://github.com/your-username/your-repository-name.git](https://github.com/your-username/your-repository-name.git)
    cd your-repository-name
    ```

2.  가상 환경 생성 및 활성화 (권장)
    ```bash
    python -m venv venv
    source venv/bin/activate  # macOS/Linux
    .\venv\Scripts\activate  # Windows
    ```

3.  의존성 설치
    ```bash
    pip install -r requirements.txt
    ```

4.  프로그램 실행
    ```bash
    python develop.py
    ```

## 📦 EXE 파일로 빌드하기

배포를 위한 단일 실행 파일은 아래 명령어로 생성할 수 있습니다.

```bash
pyinstaller --onefile --windowed --hidden-import=win32com.client --add-data "templates;templates" --add-data "업무실적계산기.xlsx;." --add-data "services;services" app.py
```
* 생성된 실행 파일은 `dist` 폴더에서 찾을 수 있습니다.

## 📝 라이선스

이 프로젝트는 [MIT License](LICENSE)를 따릅니다.