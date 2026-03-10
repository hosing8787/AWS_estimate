# AWS Excel Quote Engine (Standard Template Version)

이 프로젝트는 AWS 비용 계산기(Pricing Calculator)에서 추출된 원시 데이터(Exported Estimate CSV)를 읽어들여, 지정된 사내 공식 견적서 양식 포맷으로 100% 동일하게 오프라인 자동 변환해 주는 파이썬 기반 데스크톱/웹 애플리케이션 및 스크립트 모음입니다.

## 주요 기능
- **100% 레이아웃 호환성**: 병합된 셀, 테두리, 수식, 색상, 그리고 **기업 로고** 까지 템플릿 원본과 동일하게 렌더링.
- **동적 행 생성 (Dynamic Row Generation)**: 입력된 AWS 리소스 수량에 맞춰 항목을 매핑하고 필요한 만큼 엑셀 표의 행(Row)을 유동적으로 늘립니다. 높이 자동 조절 기능 적용.
- **다중 GUI 지원**: 사용자 친화적인 `streamlit` 기반 웹 UI 창(`app.py`) 및 `tkinter` 기반 데스크톱 UI 창(`tkinter_app.py`)을 제공하여 원클릭 변환 가능.
- **오프라인 동작**: 다운로드 받은 CSV 파일을 통신 없이 로컬에서 즉시 파싱하여 안전하게 처리.

## 필수 시스템 요구사항
- **OS**: Windows 10/11 권장
- **Python**: Python 3.8 이상
- 메모리 최소 요구사항: 4GB RAM 

## 설치 방법 (Installation)
1. Python 최신 버전이 설치되어 있는지 확인합니다. (`python --version`)
2. 이 저장소를 로컬 PC로 클론(Clone)하거나 다운로드(ZIP) 받습니다.
3. 터미널(CMD/PowerShell)을 열어 프로젝트 폴더로 이동한 후, 아래 명령어를 실행하여 필수 라이브러리를 설치합니다.
   ```bash
   pip install -r requirements.txt
   ```
   *참고: 핵심 엔진은 `openpyxl` 라이브러리에 의존하며 이미지 및 데이터 제어를 빌트인 모듈 방식으로 지원합니다.*

## 사용 방법 (Usage)

### 1. 옵션 A - 웹 GUI 리포트 생성기 실행 (추천)
```bash
streamlit run app.py
```
- 브라우저가 열리며 웹 인터페이스가 구동됩니다.
- AWS에서 내보낸 CSV 결과 파일을 드래그앤드롭하여 업로드합니다.
- 로드된 데이터 변환이 완료되면 직접 브라우저 상의 버튼을 눌러 결과물 엑셀을 다운로드합니다.

### 2. 옵션 B - 데스크톱 GUI 버전 실행
```bash
python tkinter_app.py
```
- 실행 시 작은 윈도우 창이 뜹니다.
- 버튼을 눌러 AWS에서 Export한 CSV 파일을 선택합니다.
- 변환이 완료되면 새로운 엑셀 파일이 지정 폴더(`outputs/Converted_Quote_Result.xlsx`)에 즉시 저장됩니다.

*(주의: 변환 중이거나 방금 저장된 동일한 이름의 결과물 엑셀 파일을 다른 창에서 무단으로 열어둔 경우 `Permission Denied` 에러가 뜰 수 있으니, 템플릿과 결과 파일은 미리 닫고 실행하세요.)*

### 3. 코어 스크립트 (CLI) 버전 실행
UI 없이 코드/스크립트만으로 대량 처리해야 하는 경우:
```bash
python run_existing_converter.py
```
- 내부 소스 코드상에 정의된 로직과 입력 데이터를 읽어 즉시 결과물 엑셀 파일로 추출합니다. `run_existing_converter.py` 내의 파일 경로를 수정하여 사용하세요.

## 핵심 구조 알아보기
- `generate_quote.py`: 엑셀 렌더링, 색상, 레이아웃 병합 등 GUI/디자인 관련 최종 조립.
- `shared_logic.py`, `app.py`: 데이터를 파싱하고 분류, 계산하는 코어 비즈니스 로직.
- `tkinter_app.py`: 데스크톱 환경의 일반 사용자를 위한 탐색기 연동 UI.
- `template.xlsx`: 엔진이 기준으로 삼는 원본 템플릿 파일.

## Troubleshooting 팁
- **엑셀 파일을 열었을 때 복구 팝업이 뜰 경우**: 원본 템플릿(`template.xlsx`)의 XML 구조를 임의로 변경했기 때문일 수 있습니다. 정해진 원본 파일을 그대로 사용해 주시고, 엑셀 편집기로 템플릿을 무단 수정하지 말고 테스트해 주세요.
- **레이아웃이나 병합된 셀이 깨질 경우**: 현재 프로젝트에서 호환성이 보장된 `openpyxl==3.1.5` 버전을 사용하고 있는지 `pip list`로 확인해 주세요. 다른 버전 사용 시 템플릿 복원이 원활히 이루어지지 않을 수 있습니다.
