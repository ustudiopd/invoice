# Excel + GPT 분석기

엑셀 견적서를 분석하고 GPT를 통해 질의응답할 수 있는 데스크톱 애플리케이션입니다.

## 프로젝트 구조

```
InvoiceSheet/
├── src/
│   ├── main.py                # 프로그램 진입점
│   ├── config/
│   │   └── settings.py        # 환경변수 및 설정
│   ├── ui/
│   │   ├── main_window.py     # 메인 윈도우
│   │   └── widgets/
│   │       ├── zoomable_table.py
│   │       └── border_delegate.py
│   ├── services/
│   │   ├── excel_service.py   # 엑셀 처리
│   │   └── gpt_service.py     # GPT API 호출
│   └── utils/
│       └── color_utils.py     # 색상 유틸리티
├── requirements.txt           # 프로젝트 의존성
├── .env                       # 환경 변수 설정 파일
├── .gitignore                 # Git 무시 파일 목록
```

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
  - DROPBOX 관련 설정 (필요시)
  - CHATGPT_API_KEY: OpenAI API 키
  - CHATGPT_MODEL: 사용할 GPT 모델 (기본값: gpt-4.1-mini)

3. 애플리케이션 실행:
**반드시 프로젝트 루트(InvoiceSheet)에서 아래 명령어로 실행하세요:**
```bash
python -m src.main
```

## 코드 컨벤션

- 단일 책임 원칙에 따라 모듈화
- 클래스: PascalCase, 함수/변수: snake_case
- 들여쓰기 4칸, 최대 줄 길이 79자
- 예외는 구체적으로 처리, 사용자에게 명확한 메시지 제공
- 민감 정보는 .env 파일에 저장, .gitignore에 포함

## 기타

- 실행 중 오류가 발생하면 콘솔 로그와 하단 로그창을 참고하세요.
- 개선사항/버그는 이슈로 등록해 주세요. 