"""
메인 실행 모듈 (Main Entry Point)

한화큐셀 번역 프로젝트의 메인 실행 파일입니다.
Google Sheets에서 작업을 가져와 순차적으로 처리합니다.
"""

import os
import shutil
import traceback

from docx import Document
from pptx import Presentation

from .config import (
    ORIGIN_FOLDER,
    COMPLETED_FOLDER,
    SUPPORTED_EXTENSIONS,
    ALL_SUPPORTED_EXTENSIONS,
    validate_config
)
from .translator import generate_context
from .handlers import process_docx, process_pptx, process_xlsx
from .sheets_manager import SheetsManager, Status
from .converter import convert_doc_to_docx, needs_conversion, get_converted_extension
from .slack_notifier import send_completion_notification, send_error_notification
from .glossary import get_glossary


def normalize_extension(file_name):
    """
    파일명의 확장자를 소문자로 변환합니다.
    
    Args:
        file_name (str): 파일명 (확장자 포함)
        
    Returns:
        tuple: (정규화된 파일명, 원본 확장자가 대문자였는지 여부)
    """
    name, ext = os.path.splitext(file_name)
    ext_lower = ext.lower()
    
    # 확장자가 대문자였는지 확인
    was_uppercase = (ext != ext_lower)
    
    # 소문자로 변환된 파일명
    normalized_name = f"{name}{ext_lower}"
    
    return normalized_name, was_uppercase


def build_file_path(upper_path, sub_path, file_name):
    """
    상위경로, 세부경로, 파일명을 조합하여 전체 파일 경로를 생성합니다.
    
    .doc 파일의 경우 작업 파일은 .docx로 생성됩니다.
    확장자가 대문자인 경우 소문자로 변환합니다.
    
    Args:
        upper_path (str): 상위 경로 (예: "MC")
        sub_path (str): 세부 경로 (예: "10.분석단계")
        file_name (str): 파일명 (확장자 포함)
        
    Returns:
        tuple: (원본 파일 경로, 완료 폴더 경로, 원본 복사본 경로, 작업 파일 경로, 정규화된 파일명)
    """
    # 원본 파일 경로 (원본 파일명 그대로 사용)
    origin_path = os.path.join(ORIGIN_FOLDER, upper_path, sub_path, file_name)
    
    # 완료 폴더 내 경로 (동일 구조 유지)
    completed_dir = os.path.join(COMPLETED_FOLDER, upper_path, sub_path)
    
    # 확장자 소문자로 정규화
    normalized_file_name, _ = normalize_extension(file_name)
    
    # 원본 복사본 경로 (확장자 소문자로 저장)
    completed_original = os.path.join(completed_dir, normalized_file_name)
    
    # 작업 파일 경로 결정
    name, ext = os.path.splitext(normalized_file_name)
    
    # .doc 파일은 .docx로 변환하여 작업
    if ext.lower() == '.doc':
        work_file_name = f"{name} - en.docx"  # .doc → .docx 변환
    else:
        work_file_name = f"{name} - en{ext}"
    
    work_file_path = os.path.join(completed_dir, work_file_name)
    
    return origin_path, completed_dir, completed_original, work_file_path, normalized_file_name


def prepare_work_files(origin_path, completed_dir, completed_original, work_file_path):
    """
    작업 파일을 준비합니다.
    - 완료 폴더에 동일 경로 생성
    - 원본 파일 복사 (백업)
    - 작업 파일 생성 (번역 대상)
    - .doc 파일은 .docx로 변환
    
    Args:
        origin_path (str): 원본 파일 경로
        completed_dir (str): 완료 폴더 경로
        completed_original (str): 원본 복사본 경로
        work_file_path (str): 작업 파일 경로
        
    Returns:
        str: 실제 작업 파일 경로 (성공 시)
        None: 실패 시
    """
    try:
        # 1. 완료 폴더에 동일 경로 생성
        os.makedirs(completed_dir, exist_ok=True)
        
        # 2. 원본 파일 복사 (백업용 - 원본 형식 그대로)
        if not os.path.exists(completed_original):
            shutil.copy2(origin_path, completed_original)
            print(f"   📁 원본 복사 완료: {os.path.basename(completed_original)}")
        
        # 3. 파일 형식에 따른 작업 파일 생성
        ext = os.path.splitext(origin_path)[1].lower()
        
        if ext == '.doc':
            # .doc → .docx 변환
            # 먼저 원본을 임시로 복사한 후 변환
            temp_doc_path = os.path.join(completed_dir, os.path.basename(origin_path))
            if not os.path.exists(temp_doc_path):
                shutil.copy2(origin_path, temp_doc_path)
            
            # .docx로 변환
            convert_doc_to_docx(temp_doc_path, work_file_path)
            
            # 임시 .doc 파일 삭제 (원본 복사본이 이미 있으므로)
            if temp_doc_path != completed_original:
                try:
                    os.remove(temp_doc_path)
                except:
                    pass
                    
            print(f"   📝 작업 파일 생성 (변환됨): {os.path.basename(work_file_path)}")
        else:
            # 다른 형식은 그대로 복사
            shutil.copy2(origin_path, work_file_path)
            print(f"   📝 작업 파일 생성: {os.path.basename(work_file_path)}")
        
        return work_file_path
        
    except Exception as e:
        print(f"   ❌ 파일 준비 실패: {e}")
        return None


def extract_sample_text(file_path):
    """
    파일에서 Context 분석용 샘플 텍스트를 추출합니다.
    
    Args:
        file_path (str): 파일 경로
        
    Returns:
        str: 추출된 샘플 텍스트
    """
    sample_text = ""
    
    # 확장자를 소문자로 변환하여 대소문자 구분 없이 처리
    file_path_lower = file_path.lower()
    
    if file_path_lower.endswith('.docx'):
        doc = Document(file_path)
        sample_text = "\n".join([p.text for p in doc.paragraphs[:300]])
        
    elif file_path_lower.endswith('.pptx'):
        prs = Presentation(file_path)
        for i, slide in enumerate(prs.slides):
            if i >= 3:
                break
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    sample_text += shape.text + "\n"
                    
    elif file_path_lower.endswith('.xlsx'):
        sample_text = "MES Excel Data"
    
    return sample_text


def process_single_file(work_file_path, file_context, sheets_manager=None, row_index=None):
    """
    단일 파일을 번역 처리합니다.
    
    Args:
        work_file_path (str): 작업 파일 경로
        file_context (str): 번역 지침 (Context)
        sheets_manager (SheetsManager, optional): 시트 관리자 (진행 상황 추적용)
        row_index (int, optional): 시트 행 번호
        
    Returns:
        str: 번역된 파일 경로 (성공 시)
        None: 실패 시
    """
    # 확장자를 소문자로 변환하여 대소문자 구분 없이 처리
    file_path_lower = work_file_path.lower()
    
    if file_path_lower.endswith('.docx'):
        return process_docx(work_file_path, file_context, sheets_manager, row_index)
    elif file_path_lower.endswith('.pptx'):
        return process_pptx(work_file_path, file_context, sheets_manager, row_index)
    elif file_path_lower.endswith('.xlsx'):
        return process_xlsx(work_file_path, file_context, sheets_manager, row_index)
    else:
        print(f"   ⚠️ 지원하지 않는 파일 형식: {work_file_path}")
        return None


def process_task(sheets_manager, task):
    """
    단일 작업을 처리합니다.
    
    - 상태가 "대기"인 경우: 파일 복사 후 번역 시작
    - 상태가 "진행중" 또는 "오류"인 경우: 기존 "-en" 파일로 이어서 번역
    
    Args:
        sheets_manager (SheetsManager): 시트 관리자
        task (dict): 작업 정보
        
    Returns:
        bool: 성공 여부
    """
    row_index = task['row_index']
    upper_path = task['upper_path']
    sub_path = task['sub_path']
    file_name = task['file_name']
    current_status = task.get('status', '대기')  # 현재 상태
    
    # 이어하기 모드 여부
    is_resume_mode = current_status in ['진행중', '오류']
    
    print(f"\n{'='*60}")
    print(f"📄 파일 처리 시작: {file_name}")
    print(f"   경로: {upper_path}/{sub_path}")
    if is_resume_mode:
        print(f"   🔄 이어하기 모드 (이전 상태: {current_status})")
    print(f"{'='*60}")
    
    try:
        # 1. 파일 경로 구성 (확장자 소문자로 정규화)
        origin_path, completed_dir, completed_original, work_file_path, normalized_file_name = build_file_path(
            upper_path, sub_path, file_name
        )
        
        # 2. 확장자가 대문자였으면 Google Sheets에서 파일명 업데이트
        if file_name != normalized_file_name:
            print(f"   📝 파일명 확장자 정규화: {file_name} → {normalized_file_name}")
            sheets_manager.update_file_name(row_index, normalized_file_name)
            file_name = normalized_file_name  # 이후 로직에서 사용할 파일명 업데이트
        
        # 3. 원본 파일 존재 확인
        if not os.path.exists(origin_path):
            raise FileNotFoundError(f"원본 파일을 찾을 수 없습니다: {origin_path}")
        
        # 4. 파일 확장자 확인 (최신 형식 + 변환 가능 형식 모두 허용)
        ext = os.path.splitext(file_name)[1].lower()
        if ext not in ALL_SUPPORTED_EXTENSIONS:
            raise ValueError(f"지원하지 않는 파일 형식입니다: {ext}")
        
        # 5. 원본 파일 무결성 검사 (손상된 파일 조기 감지)
        if not verify_file_integrity(origin_path):
            raise ValueError(f"원본 파일이 손상되었습니다 (열 수 없음): {origin_path}")
        
        # 5. 진행상태 '진행중'으로 변경 + 시작시간 기록
        # 이어하기 모드에서는 토큰 초기화 하지 않음
        if is_resume_mode:
            sheets_manager.update_status(row_index, '진행중')
            print(f"   ✅ 상태 변경: 진행중 (이어하기)")
        else:
            sheets_manager.start_task(row_index)
            print(f"   ✅ 상태 변경: 진행중")
        
        # 6. 작업 파일 준비
        if is_resume_mode:
            # 이어하기 모드: 기존 "-en" 파일 사용
            actual_work_path = prepare_work_files_resume(work_file_path, origin_path, completed_dir, completed_original)
        else:
            # 새 작업: 파일 복사 후 시작
            actual_work_path = prepare_work_files(origin_path, completed_dir, completed_original, work_file_path)
        
        if not actual_work_path:
            raise Exception("작업 파일 준비 실패")
        
        # 7. Context 분석
        print(f"   🤖 Context 분석 중...")
        sample_text = extract_sample_text(actual_work_path)
        file_context = generate_context(sample_text)
        print(f"   ✅ Context 분석 완료")
        
        # 8. 번역 실행 (시트 진행 상황 추적 포함)
        result = process_single_file(actual_work_path, file_context, sheets_manager, row_index)
        
        if result:
            # 9. 완료 처리
            sheets_manager.mark_completed(row_index)
            print(f"\n   🎉 번역 완료!")
            
            # 10. Slack 완료 알림 전송
            try:
                times = sheets_manager.get_task_times(row_index)
                progress = sheets_manager.get_overall_progress()
                file_path = f"{upper_path}/{sub_path}"
                
                send_completion_notification(
                    file_name=file_name,
                    file_path=file_path,
                    start_time=times['start_time'],
                    end_time=times['end_time'],
                    progress_percent=progress
                )
            except Exception as slack_error:
                print(f"   ⚠️ Slack 알림 전송 실패: {slack_error}")
            
            return True
        else:
            raise Exception("번역 처리 실패")
            
    except Exception as e:
        # 오류 발생 시 기록
        error_msg = str(e)
        error_trace = traceback.format_exc()
        module_name = "main.process_task"
        
        # 상세 오류 메시지 구성
        detailed_error = f"{error_msg}\n\n상세:\n{error_trace}"
        
        print(f"\n   ❌ 오류 발생: {error_msg}")
        sheets_manager.record_error(row_index, detailed_error, module_name)
        
        # Slack 오류 알림 전송
        try:
            slack_error_msg = f"*파일*: {file_name}\n*경로*: {upper_path}/{sub_path}\n*오류*: {error_msg}"
            send_error_notification(slack_error_msg)
        except Exception as slack_error:
            print(f"   ⚠️ Slack 알림 전송 실패: {slack_error}")
        
        return False


def verify_file_integrity(file_path):
    """
    파일의 무결성을 검사합니다.
    
    손상된 파일(빈 파일, 열리지 않는 파일)인지 확인합니다.
    
    Args:
        file_path (str): 검사할 파일 경로
        
    Returns:
        bool: 파일이 정상이면 True, 손상되었으면 False
    """
    try:
        # 파일 크기 확인 (0바이트면 손상)
        if os.path.getsize(file_path) == 0:
            return False
        
        # 확장자별 무결성 검사
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext == '.docx':
            # python-docx로 열어보기
            doc = Document(file_path)
            # 최소한 body가 있는지 확인
            _ = doc.paragraphs
            return True
            
        elif ext == '.pptx':
            # python-pptx로 열어보기
            prs = Presentation(file_path)
            _ = prs.slides
            return True
            
        elif ext == '.xlsx':
            # 파일 크기만 확인 (xlwings는 Excel 필요)
            return os.path.getsize(file_path) > 1000  # 최소 1KB
            
        return True  # 다른 형식은 기본적으로 통과
        
    except Exception as e:
        print(f"   ⚠️ 파일 무결성 검사 실패: {e}")
        return False


def prepare_work_files_resume(work_file_path, origin_path, completed_dir, completed_original):
    """
    이어하기 모드에서 작업 파일을 준비합니다.
    
    기존 "-en" 파일이 있으면 무결성 검사 후 사용하고,
    손상되었거나 없으면 새로 생성합니다.
    
    Args:
        work_file_path (str): 작업 파일 경로 ("-en" 파일)
        origin_path (str): 원본 파일 경로
        completed_dir (str): 완료 폴더 경로
        completed_original (str): 원본 복사본 경로
        
    Returns:
        str: 실제 작업 파일 경로 (성공 시)
        None: 실패 시
    """
    try:
        # 1. "-en" 파일이 이미 존재하는지 확인
        if os.path.exists(work_file_path):
            print(f"   ✅ 기존 작업 파일 발견: {os.path.basename(work_file_path)}")
            
            # 2. 파일 무결성 검사
            if verify_file_integrity(work_file_path):
                print(f"   🔄 이어서 번역을 진행합니다...")
                return work_file_path
            else:
                # 손상된 파일 삭제 후 새로 생성
                print(f"   ⚠️ 기존 파일이 손상되었습니다. 삭제 후 새로 생성합니다...")
                try:
                    os.remove(work_file_path)
                except Exception as del_err:
                    print(f"   ❌ 손상된 파일 삭제 실패: {del_err}")
                    return None
        
        # 3. "-en" 파일이 없거나 손상되었으면 새로 생성
        print(f"   📝 작업 파일을 새로 생성합니다...")
        return prepare_work_files(origin_path, completed_dir, completed_original, work_file_path)
        
    except Exception as e:
        print(f"   ❌ 이어하기 파일 준비 실패: {e}")
        return None


def main():
    """
    메인 함수 - Google Sheets 기반 작업 처리의 진입점입니다.
    """
    print("=" * 60)
    print("🌐 한화큐셀 번역 프로젝트 v1.3.0")
    print("   Google Sheets 연동 대량 처리 모드")
    print("   .doc → .docx 자동 변환 지원")
    print("   📚 용어집(Glossary) 프롬프트 주입 지원")
    print("=" * 60)
    
    # 1. 설정 검증
    is_valid, message = validate_config()
    if not is_valid:
        print(f"\n❌ 설정 오류: {message}")
        return
    
    print("\n✅ 설정 검증 완료")
    
    # 2. 용어집 로드
    glossary = get_glossary()
    if glossary.is_loaded:
        print(f"✅ 용어집 로드 완료: {glossary.get_term_count()}개 용어")
    else:
        print("⚠️ 용어집 없이 진행합니다 (data/용어정의.xlsx 파일 확인 필요)")
    
    # 3. Google Sheets 연결
    try:
        sheets_manager = SheetsManager()
    except Exception as e:
        print(f"\n❌ Google Sheets 연결 실패: {e}")
        print("\n💡 해결 방법:")
        print("   1. credentials.json 파일이 프로젝트 루트에 있는지 확인하세요.")
        print("   2. 서비스 계정에 스프레드시트 접근 권한이 있는지 확인하세요.")
        return
    
    # 4. 작업 루프 시작
    success_count = 0
    fail_count = 0
    
    # 연속 오류 방지를 위한 추적 변수 (파일명 기반)
    fail_count_by_file = {}  # {file_name: 실패 횟수}
    MAX_CONSECUTIVE_FAILS = 3  # 같은 파일 연속 실패 허용 횟수
    skipped_files = set()  # 건너뛴 파일명 목록
    
    print("\n🚀 작업 시작...")
    print("   (Ctrl+C로 중단할 수 있습니다)")
    
    while True:
        try:
            # 대기 중인 작업 조회
            task = sheets_manager.get_next_waiting_task()
            
            if task is None:
                print("\n" + "=" * 60)
                print("✅ 모든 대기 작업 완료!")
                break
            
            current_row = task['row_index']
            file_name = task['file_name']
            
            # 이미 건너뛴 파일이면 "건너뛰기완료"로 변경하고 다음으로
            if file_name in skipped_files:
                print(f"\n⏭️ 건너뛰기: {file_name} (반복 실패로 제외됨)")
                sheets_manager.update_status(current_row, "건너뛰기완료")
                continue
            
            # 해당 파일의 실패 횟수 확인
            current_fail_count = fail_count_by_file.get(file_name, 0)
            
            if current_fail_count >= MAX_CONSECUTIVE_FAILS:
                print(f"\n⚠️ 파일 '{file_name}'이(가) {MAX_CONSECUTIVE_FAILS}회 실패했습니다.")
                print(f"   → 이 파일을 건너뛰고 다음 파일로 이동합니다.")
                skipped_files.add(file_name)
                sheets_manager.update_status(current_row, "건너뛰기완료")
                fail_count += 1
                continue
            
            # 작업 처리
            if process_task(sheets_manager, task):
                success_count += 1
                # 성공 시 실패 카운트 제거
                if file_name in fail_count_by_file:
                    del fail_count_by_file[file_name]
            else:
                fail_count += 1
                # 실패 시 카운트 증가
                fail_count_by_file[file_name] = current_fail_count + 1
                print(f"   ⚠️ 실패 횟수: {fail_count_by_file[file_name]}/{MAX_CONSECUTIVE_FAILS}")
                
        except KeyboardInterrupt:
            print("\n\n⚠️ 사용자에 의해 중단되었습니다.")
            break
        except Exception as e:
            print(f"\n❌ 예기치 않은 오류: {e}")
            fail_count += 1
    
    # 5. 최종 결과 출력
    print("\n" + "=" * 60)
    print("📊 작업 완료 요약")
    print("=" * 60)
    print(f"   ✅ 성공: {success_count}개")
    print(f"   ❌ 실패: {fail_count}개")
    print(f"   📁 결과 위치: {COMPLETED_FOLDER}")
    print("=" * 60)


if __name__ == "__main__":
    main()
