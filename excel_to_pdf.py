import os
import win32com.client
import re
import traceback
import time


def safe_filename(filename):
    # 파일명을 그대로 유지 (변환 없음)
    return filename


def excel_to_pdf(excel_path, pdf_path=None):
    print("\n=== PDF 변환 시작 ===")
    print(f"1. 입력 파일: {excel_path}")
    print(f"2. 출력 경로: {pdf_path}")
    print(f"3. 파일 존재 여부:")
    print(f"   - 엑셀 파일: {os.path.exists(excel_path)}")
    print(f"   - PDF 폴더: {os.path.exists(os.path.dirname(pdf_path))}")
    print(f"4. 파일 권한:")
    print(f"   - 엑셀 읽기: {os.access(excel_path, os.R_OK)}")
    print(f"   - PDF 쓰기: {os.access(os.path.dirname(pdf_path), os.W_OK)}")
    
    # PDF 파일이 이미 존재하면 삭제
    if os.path.exists(pdf_path):
        try:
            os.remove(pdf_path)
            print(f"   - 기존 PDF 파일 삭제: {pdf_path}")
        except Exception as e:
            print(f"   - 기존 PDF 파일 삭제 실패: {e}")
    
    try:
        print("\n5. Excel 애플리케이션 시작...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False  # 경고창 표시 안 함
        print("   - Excel 애플리케이션 생성 성공")
        
        print("\n6. 워크북 열기...")
        abs_excel_path = os.path.abspath(excel_path)
        abs_pdf_path = os.path.abspath(pdf_path)
        print(f"   - 엑셀 절대 경로: {abs_excel_path}")
        print(f"   - PDF 절대 경로: {abs_pdf_path}")
        
        wb = excel.Workbooks.Open(abs_excel_path)
        print("   - 워크북 열기 성공")
        
        print("\n7. PDF 변환 시작...")
        print(f"   - 저장 경로: {abs_pdf_path}")
        
        # PDF 변환 시도 (최대 3번)
        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                wb.ExportAsFixedFormat(0, abs_pdf_path)
                print("   - PDF 변환 성공")
                break
            except Exception as e:
                if attempt < max_attempts - 1:
                    print(f"   - PDF 변환 실패 (시도 {attempt + 1}/{max_attempts})")
                    print(f"   - 오류: {str(e)}")
                    time.sleep(1)  # 1초 대기 후 재시도
                else:
                    raise
        
    except Exception as e:
        print("\n❌ 오류 발생!")
        print(f"오류 유형: {type(e).__name__}")
        print(f"오류 메시지: {str(e)}")
        print("\n상세 오류 정보:")
        traceback.print_exc()
        raise
    finally:
        print("\n8. 정리 작업...")
        try:
            wb.Close(False)
            print("   - 워크북 닫기 성공")
        except Exception as e:
            print(f"   - 워크북 닫기 실패: {e}")
        try:
            excel.Quit()
            print("   - Excel 종료 성공")
        except Exception as e:
            print(f"   - Excel 종료 실패: {e}")
        print("=== PDF 변환 종료 ===\n")


if __name__ == "__main__":
    folder = os.path.abspath("./2025년 견적서_주식회사")
    pdf_dir = os.path.join(folder, "PDF")
    os.makedirs(pdf_dir, exist_ok=True)
    
    print(f"\n작업 폴더: {folder}")
    print(f"PDF 저장 폴더: {pdf_dir}")
    
    for fname in os.listdir(folder):
        if fname.lower().endswith((".xlsx", ".xlsm", ".xls")):
            excel_path = os.path.join(folder, fname)
            # 확장자만 .pdf로 변경
            pdf_name = os.path.splitext(fname)[0] + ".pdf"
            pdf_path = os.path.join(pdf_dir, pdf_name)
            try:
                excel_to_pdf(excel_path, pdf_path)
            except Exception as e:
                print(f"\n❌ 파일 처리 실패: {fname}")
                print(f"오류: {str(e)}")
                continue 