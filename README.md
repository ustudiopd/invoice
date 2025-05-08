# Excel + GPT 분석기

엑셀 견적서를 분석하고 GPT를 통해 질의응답할 수 있는 데스크톱 애플리케이션입니다.

## 프로젝트 구조

```
InvoiceSheet/
├── .env                    # 환경 변수 설정 파일
├── .gitignore             # Git 무시 파일 목록
├── excel_gpt_viewer.py    # 메인 애플리케이션 코드
└── requirements.txt       # 프로젝트 의존성
```

## 코드 컨벤션

### 1. 파일 구조
- 각 파일은 단일 책임 원칙을 따릅니다.
- 메인 애플리케이션 로직은 `excel_gpt_viewer.py`에 포함됩니다.

### 2. 클래스 네이밍
- 클래스 이름은 PascalCase를 사용합니다.
- 예: `ExcelGPTViewer`, `ZoomableTableWidget`, `BorderDelegate`

### 3. 함수/메서드 네이밍
- 함수와 메서드 이름은 snake_case를 사용합니다.
- 예: `open_excel()`, `ask_gpt()`, `apply_tint()`

### 4. 변수 네이밍
- 변수 이름은 snake_case를 사용합니다.
- 예: `json_path`, `excel_path`, `zoom_factor`

### 5. 상수 네이밍
- 상수는 대문자와 언더스코어를 사용합니다.
- 예: `DROPBOX_APP_KEY`, `GPT_API_KEY`

### 6. 주석 작성
- 함수와 클래스는 docstring을 사용하여 설명합니다.
- 복잡한 로직에는 인라인 주석을 추가합니다.

### 7. 코드 포맷팅
- 들여쓰기는 4칸 스페이스를 사용합니다.
- 최대 줄 길이는 79자를 초과하지 않습니다.

### 8. 예외 처리
- 모든 예외는 구체적으로 처리합니다.
- 사용자에게 적절한 오류 메시지를 표시합니다.

### 9. 환경 변수
- 민감한 정보는 `.env` 파일에 저장합니다.
- `.env` 파일은 `.gitignore`에 포함되어 있습니다.

### 10. Git 관리
- `.gitignore`에 다음 항목들이 포함됩니다:
  - `.env` 파일
  - `2025년 견적서_주식회사/` 디렉토리

## 주요 기능

1. 엑셀 파일 로드 및 표시
2. 확대/축소 기능
3. GPT를 통한 견적서 분석
4. JSON 형식 데이터 저장
5. 실시간 셀 편집 및 동기화

## 의존성

- PyQt5
- openpyxl
- python-dotenv
- requests

## 설치 및 실행

1. 의존성 설치:
```bash
pip install -r requirements.txt
```

2. `.env` 파일 설정:
- 프로젝트 루트 디렉토리에 `.env` 파일을 생성하고 다음 환경 변수들을 설정합니다:
  - DROPBOX 관련 설정 (APP_KEY, APP_SECRET, ACCESS_TOKEN, REFRESH_TOKEN, SHARED_FOLDER_ID, SHARED_FOLDER_NAME)
  - LOCAL_BID_FOLDER: 로컬 견적서 폴더 경로
  - CHATGPT_API_KEY: OpenAI API 키
  - CHATGPT_MODEL: 사용할 GPT 모델 (기본값: gpt-4.1-mini)

3. 애플리케이션 실행:
```bash
python excel_gpt_viewer.py
``` 