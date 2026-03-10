# Azure Excel Quote Engine (Standard Template Version)

이 프로젝트는 Azure 가격 계산기에서 추출된 원시 데이터(Exported Estimate Excel)를 읽어들여, 지정된 사내 공식 견적서 양식 포맷으로 100% 동일하게 오프라인 자동 변환해 주는 파이썬 기반 데스크톱 애플리케이션 및 스크립트 모음입니다.

## 주요 기능
- **100% 레이아웃 호환성**: 병합된 셀, 테두리, 수식, 색상, 그리고 **기업 로고(SK 이미지 등)** 까지 템플릿 원본과 동일하게 렌더링.
- **동적 행 생성 (Dynamic Row Generation)**: 입력된 Azure 리소스 수량에 맞춰 `Server 및 DBMS`, `기타자원`, `모니터링` 등의 카테고리로 매핑하고 필요한 만큼 엑셀 표의 행(Row)을 유동적으로 늘립니다. 높이 자동 조절 기능 적용.
- **GUI 지원**: 사용자 친화적인 `tkinter` 기반 Window UI 창을 통해 파일 선택 및 원클릭 변환 가능.
- **오프라인 동작 (Live Price API 선택적 활용)**: 다운로드 받은 Export 파일을 로컬에서 즉시 파싱. 필요 시 소스 코드 내에서 최신 Azure Retail Price API 연결도 가능.

## 필수 시스템 요구사항
- **OS**: Windows 10/11 권장
- **Python**: Python 3.8 이상 (테스트 환경: Python 3.15)
- 메모리 최소 요구사항: 4GB RAM 

## 설치 방법 (Installation)
1. Python 최신 버전이 설치되어 있는지 확인합니다. (`python --version`)
2. 이 저장소를 로컬 PC로 클론(Clone)하거나 다운로드(ZIP) 받습니다.
3. 터미널(CMD/PowerShell)을 열어 프로젝트 폴더로 이동한 후, 아래 명령어를 실행하여 필수 라이브러리를 설치합니다.
   ```bash
   pip install -r requirements.txt
   ```
   *참고: 핵심 엔진은 `openpyxl` 라이브러리에 의존하며 이미지 및 XML 포팅을 위해 빌트인 모듈(`zipfile`, `json` 등)을 사용합니다.*

## 사용 방법 (Usage)

### 1. 데스크톱 GUI 버전 실행 (추천)
```bash
python tkinter_app.py
```
- 실행 시 작은 윈도우 창이 뜹니다.
- `불러오기` 버튼을 눌러 Azure에서 Export한 Excel 파일(예: `ExportedEstimate_1.xlsx`)을 선택합니다.
- `Excel로 변환하기` 버튼을 누르면, 양식 템플릿(`.xlsx`)을 기반으로 데이터가 들어가고 병합이 완료된 새로운 엑셀 파일이 같은 폴더에 저장됩니다.

*(주의: 변환 중이거나 방금 저장된 동일한 이름의 결과물 엑셀 파일을 다른 창에서 무단으로 열어둔 경우 `Permission Denied` 에러가 뜰 수 있으니, 템플릿과 결과 파일은 미리 닫고 실행하세요.)*

### 2. 코어 스크립트 (CLI) 버전 실행
UI 없이 코드/스크립트만으로 대량 처리해야 하는 경우:
```bash
python run_existing_converter.py
```
- 내부 소스 코드상에 정의된 `inputs/` 폴더 내의 파일을 읽어 즉시 결과물 엑셀 파일로 추출합니다. `run_existing_converter.py` 내의 파일 경로를 수정하여 사용하세요.

## 핵심 구조 알아보기
- `generate_quote.py`: 엑셀 렌더링, 색상, 레이아웃 병합, 이미지 ZIP 인젝션 등 GUI/디자인 관련 최종 조립.
- `shared_logic.py`, `app.py`: 데이터를 파싱하고 분류, 계산하는 코어 비즈니스 로직.
- `tkinter_app.py`: 일반 사용자를 위한 파일 탐색기 연동 UI.

## Troubleshooting 팁
- **엑셀 파일을 열었을 때 복구 팝업이 뜰 경우**: `generate_quote.py`의 ZIP/XML 치환 로직에서 namespace(`r:id`)가 충돌했을 가능성이 있습니다. 원본 템플릿 구조를 변경하지 말고 테스트해 주세요.
- **SK 로고 이미지가 안 보일 경우**: 로고 삽입은 Python 3.15+ 환경에서 C++ 바인딩을 피하기 위해 `Pillow`를 쓰지 않고 ZIP XML 바이너리 복사 방식을 채택했습니다. 템플릿(`orig.xlsx`) 경로가 유효한지 확인하세요.
