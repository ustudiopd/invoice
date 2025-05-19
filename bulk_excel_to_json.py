import os
from src.services.excel_service import ExcelService
import json
import tempfile
import pandas as pd
import openpyxl

def process_excel_files():
    """엑셀 파일들을 처리하여 JSON으로 변환합니다."""
    excel_dir = "./2025년 견적서_주식회사"
    json_dir = os.path.join(excel_dir, "Json_1")
    
    # JSON 디렉토리가 없으면 생성
    if not os.path.exists(json_dir):
        os.makedirs(json_dir)
    
    excel_service = ExcelService()
    
    # 디렉토리 내의 모든 파일 처리
    for filename in os.listdir(excel_dir):
        if filename.endswith((".xlsx", ".xls")):
            file_path = os.path.join(excel_dir, filename)
            tmp_path = None
            try:
                # .xls 파일은 pandas+xlrd로 읽어서 임시 .xlsx로 변환
                if filename.endswith(".xls"):
                    df = pd.read_excel(file_path, engine="xlrd")
                    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
                    os.close(fd)
                    with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False)
                    file_to_open = tmp_path
                else:
                    file_to_open = file_path

                # ExcelService를 사용하여 파일 처리
                json_path, items = excel_service.load_excel(file_to_open, None)

                # 파일명 기준으로 json 저장 경로 생성
                base = os.path.splitext(filename)[0]
                out_json_path = os.path.join(json_dir, f"{base}.json")

                # 최소 JSON 구조(meta, items, header, summary)
                data = {
                    "meta": {"file_name": filename},
                    "items": items,
                    "header": {},
                    "summary": {}
                }
                with open(out_json_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                print(f"저장 완료: {out_json_path}")

            except Exception as e:
                print(f"파일 처리 실패 ({filename}): {str(e)}")
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    os.remove(tmp_path)

if __name__ == "__main__":
    process_excel_files() 