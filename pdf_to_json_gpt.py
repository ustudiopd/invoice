import os
import json
import openai
from dotenv import load_dotenv
from PyPDF2 import PdfReader
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import time
import jsonschema
from typing import Dict, Any
import re


# JSON 스키마 정의
INVOICE_SCHEMA = {
    "type": "object",
    "required": ["견적번호", "견적일자", "거래처명", "카테고리", "합계금액", "세액", "총액"],
    "properties": {
        "견적번호": {"type": "string"},
        "견적일자": {"type": "string"},
        "거래처명": {"type": "string"},
        "카테고리": {
            "type": "array",
            "items": {
                "type": "object",
                "required": ["category", "amount", "items"],
                "properties": {
                    "category": {"type": "string"},
                    "amount": {"type": ["number", "string"]},
                    "items": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "required": ["품목명", "수량", "단가", "금액"],
                            "properties": {
                                "품목명": {"type": "string"},
                                "수량": {"type": ["number", "string"]},
                                "단가": {"type": ["number", "string"]},
                                "금액": {"type": ["number", "string"]}
                            }
                        }
                    }
                }
            }
        },
        "합계금액": {"type": ["number", "string"]},
        "세액": {"type": ["number", "string"]},
        "총액": {"type": ["number", "string"]}
    }
}


def validate_json(json_data: Dict[str, Any]) -> bool:
    """JSON 데이터가 스키마를 준수하는지 검증"""
    try:
        jsonschema.validate(instance=json_data, schema=INVOICE_SCHEMA)
        return True
    except jsonschema.exceptions.ValidationError as e:
        print(f"JSON 검증 실패: {str(e)}")
        return False


def extract_text_from_pdf(pdf_path: str) -> str:
    """PDF 파일에서 텍스트 추출"""
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
        print(f"\n[디버깅] PDF 텍스트 추출 결과 ({os.path.basename(pdf_path)}):")
        print(f"추출된 텍스트 길이: {len(text)} 문자")
        print(f"텍스트 미리보기: {text[:500]}...")
        return text
    except Exception as e:
        print(f"PDF 텍스트 추출 실패 ({pdf_path}): {str(e)}")
        return ""


def replace_none_with_empty(obj):
    """재귀적으로 None 값을 모두 공란("")으로 변환"""
    if isinstance(obj, dict):
        return {k: replace_none_with_empty(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [replace_none_with_empty(v) for v in obj]
    elif obj is None:
        return ""
    else:
        return obj


def split_text(text, max_length=3000):
    """텍스트를 max_length 단위로 분할"""
    lines = text.splitlines(keepends=True)
    chunks = []
    current = ''
    for line in lines:
        if len(current) + len(line) > max_length:
            chunks.append(current)
            current = ''
        current += line
    if current:
        chunks.append(current)
    return chunks


def gpt_extract_json(
    text: str,
    model: str = "gpt-4.1-mini",
    max_retries: int = 3
) -> Dict[str, Any]:
    """텍스트가 길면 분할하여 여러 번 GPT에 요청, 카테고리별로 합침"""
    text_chunks = split_text(text, max_length=3000)
    merged_result = None
    for idx, chunk in enumerate(text_chunks):
        print(f"[디버깅] GPT 분할 요청 {idx+1}/{len(text_chunks)} (chunk 길이: {len(chunk)})")
        prompt = (
            f"""
아래는 견적서 PDF에서 추출한 텍스트입니다.\n
이 텍스트를 아래 JSON 스키마에 맞게 변환해 주세요.

- 파란색 행(혹은 대문자/굵은 글씨 등 카테고리로 보이는 행)은 'category'로, 그 아래 하얀색 행들은 해당 카테고리의 'items'로 묶어주세요.
- 각 카테고리의 'amount'는 해당 items의 '금액' 합계로 자동 계산해 주세요.
- 표 맨 아래의 TOTAL, Tax, TOTAL Due 등은 각각 '합계금액', '세액', '총액'에 넣어주세요.

스키마:
{json.dumps(INVOICE_SCHEMA, ensure_ascii=False, indent=2)}

PDF 텍스트:
{chunk}

반드시 위 스키마에 맞는 JSON만 출력해 주세요.
숫자는 가능한 숫자 타입으로 변환해주세요 (예: "1,000" -> 1000).
"""
        )
        for attempt in range(max_retries):
            try:
                print(f"\n[디버깅] OpenAI API 호출 시도 {attempt+1}/{max_retries}")
                print(f"[디버깅] 모델명: {model}")
                print(f"[디버깅] 프롬프트 일부: {prompt[:300]} ... (생략)")
                response = openai.ChatCompletion.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.0,
                    max_tokens=2048
                )
                print(f"[디버깅] OpenAI 응답 전체: {response}")
                json_str = response.choices[0].message.content.strip()
                print(f"[디버깅] OpenAI 응답 content: {json_str[:300]} ... (생략)")
                # 코드블록 제거
                if json_str.startswith("```"):
                    json_str = re.sub(r"^```[a-zA-Z]*\n?", "", json_str)
                    json_str = re.sub(r"```$", "", json_str)
                    json_str = json_str.strip()
                json_data = json.loads(json_str)
                json_data = replace_none_with_empty(json_data)
                if not merged_result:
                    merged_result = json_data
                else:
                    # 카테고리만 합침
                    if "카테고리" in merged_result and "카테고리" in json_data:
                        merged_result["카테고리"] += json_data["카테고리"]
                break
            except Exception as e:
                print(f"[디버깅] 예외 발생 (시도 {attempt+1}/{max_retries}): {e}")
                import traceback
                traceback.print_exc()
                time.sleep(1)
        else:
            raise Exception("최대 재시도 횟수 초과 (분할 요청)")
    if merged_result and validate_json(merged_result):
        print("[디버깅] JSON 스키마 검증 성공 (분할 병합)")
        return merged_result
    else:
        raise Exception("최종 JSON 스키마 검증 실패 (분할 병합)")


def process_pdf(pdf_path: str, output_dir: str) -> bool:
    """단일 PDF 파일 처리"""
    try:
        print(f"\n[디버깅] 파일 처리 시작: {os.path.basename(pdf_path)}")
        # PDF에서 텍스트 추출
        text = extract_text_from_pdf(pdf_path)
        if not text:
            print(f"[디버깅] 텍스트 추출 실패: {os.path.basename(pdf_path)}")
            return False
        print(f"[디버깅] 텍스트 추출 성공: {os.path.basename(pdf_path)}")
        # GPT로 JSON 추출
        json_data = gpt_extract_json(text)
        # JSON 파일 저장
        json_path = os.path.join(
            output_dir,
            os.path.splitext(os.path.basename(pdf_path))[0] + ".json"
        )
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        print(f"[디버깅] 파일 처리 성공: {os.path.basename(pdf_path)}")
        return True
    except Exception as e:
        print(f"[디버깅] 파일 처리 실패 ({os.path.basename(pdf_path)}): {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def main():
    # 환경 변수 로드
    load_dotenv()
    openai.api_key = os.getenv("OPENAI_API_KEY") or \
        os.getenv("CHATGPT_API_KEY")
    if not openai.api_key:
        raise ValueError(
            "OPENAI_API_KEY 또는 CHATGPT_API_KEY가 설정되지 않았습니다."
        )
    # 경로 설정
    base_dir = "./2025년 견적서_주식회사"
    pdf_dir = os.path.join(base_dir, "PDF")
    json_dir = os.path.join(base_dir, "PDFtoJSON")
    os.makedirs(json_dir, exist_ok=True)
    # PDF 폴더 내 모든 PDF 파일 자동 탐색
    pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
    print(
        f"총 {len(pdf_files)}개의 PDF 파일을 처리합니다..."
    )
    # 병렬 처리
    success_count = 0
    failed_files = []  # 실패한 파일 목록 저장
    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = {
            executor.submit(
                process_pdf,
                os.path.join(pdf_dir, pdf_file),
                json_dir
            ): pdf_file for pdf_file in pdf_files
        }
        # 진행 상황 표시
        with tqdm(total=len(pdf_files), desc="PDF 처리 중") as pbar:
            for future in as_completed(futures):
                pdf_file = futures[future]
                try:
                    if future.result():
                        success_count += 1
                    else:
                        failed_files.append(pdf_file)
                except Exception as e:
                    print(f"\n처리 실패 ({pdf_file}): {str(e)}")
                    failed_files.append(pdf_file)
                pbar.update(1)
    # 결과 출력
    print(f"\n처리 완료: {success_count}/{len(pdf_files)} 성공")
    if failed_files:
        print("\n실패한 파일 목록:")
        for file in failed_files:
            print(f"- {file}")


if __name__ == "__main__":
    main() 