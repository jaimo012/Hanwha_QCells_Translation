"""
Google Sheets 관리 모듈 (Sheets Manager)

Google Sheets API를 사용하여 작업 상태를 관리합니다.
- 대기 중인 파일 조회
- 진행 상태 업데이트
- 분석 진행 상황 추적 (분석대상/분석완료/토큰 사용량)
- 오류 기록
"""

import os
import time
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

from .config import PROJECT_ROOT, GOOGLE_SHEETS_URL, GOOGLE_SHEETS_NAME


# ==============================================================================
# [Rate Limit 설정]
# ==============================================================================
SHEETS_API_RETRY_COUNT = 3      # 재시도 횟수
SHEETS_API_RETRY_DELAY = 5      # 재시도 대기 시간 (초)
SHEETS_API_MIN_DELAY = 0.5      # API 호출 간 최소 대기 시간 (초)


# ==============================================================================
# [상수 정의] Google Sheets 컬럼 인덱스 (1-based)
# ==============================================================================
class SheetColumns:
    """시트 컬럼 인덱스 (1부터 시작)"""
    ROW_NUM = 1          # A: 연번
    UPPER_PATH = 2       # B: 상위경로
    SUB_PATH = 3         # C: 세부경로
    FILE_NAME = 4        # D: 파일명
    FILE_SIZE = 5        # E: 파일용량
    FILE_TYPE = 6        # F: 파일유형
    STATUS = 7           # G: 진행상태
    START_TIME = 8       # H: 시작일시
    END_TIME = 9         # I: 종료일시
    ERROR = 10           # J: 오류
    NOTE = 11            # K: 비고
    INPUT_TOKEN = 12     # L: 인풋토큰
    OUTPUT_TOKEN = 13    # M: 아웃풋토큰
    TOTAL_COST = 14      # N: 총비용


# ==============================================================================
# [진행 상태 값]
# ==============================================================================
class Status:
    """진행상태 값"""
    WAITING = "대기"
    IN_PROGRESS = "진행중"
    COMPLETED = "완료"
    ERROR = "오류"
    VERIFIED_1 = "1차검증완료"
    REVIEW_1_COMPLETED = "1차 검수완료"  # 1차 번역 검수 완료


# ==============================================================================
# [Google Sheets 클라이언트 클래스]
# ==============================================================================
class SheetsManager:
    """
    Google Sheets 연동 관리자
    
    작업 대기열 조회, 상태 업데이트, 오류 기록 등을 담당합니다.
    """
    
    def __init__(self):
        """SheetsManager 초기화 - Google Sheets 연결"""
        self.credentials_path = os.path.join(PROJECT_ROOT, "credentials.json")
        self.sheet = None
        self._connect()
    
    def _connect(self):
        """Google Sheets에 연결합니다."""
        try:
            # 서비스 계정 인증
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            
            creds = Credentials.from_service_account_file(
                self.credentials_path, 
                scopes=scopes
            )
            
            client = gspread.authorize(creds)
            
            # 스프레드시트 열기
            spreadsheet = client.open_by_url(GOOGLE_SHEETS_URL)
            self.sheet = spreadsheet.worksheet(GOOGLE_SHEETS_NAME)
            
            print("✅ Google Sheets 연결 성공")
            
        except Exception as e:
            print(f"❌ Google Sheets 연결 실패: {e}")
            raise
    
    def get_next_waiting_task(self):
        """
        맨 위에서부터 순서대로 '완료'가 아닌 첫 번째 행을 찾습니다.
        
        시트에 이미 우선순위대로 정렬되어 있으므로, 
        맨 위에서부터 순차적으로 처리합니다.
        
        처리 대상 상태: 대기, 진행중, 오류
        제외 상태: 완료
        
        Returns:
            dict: 작업 정보 딕셔너리 (row_index, upper_path, sub_path, file_name, status 등)
            None: 처리할 작업이 없을 때
        """
        try:
            # 모든 데이터 가져오기
            all_values = self.sheet.get_all_values()
            
            # 처리 대상 상태 목록
            processable_statuses = [Status.WAITING, Status.IN_PROGRESS, Status.ERROR]
            
            # 헤더 제외하고 맨 위에서부터 순차적으로 순회
            for idx, row in enumerate(all_values[1:], start=2):  # 실제 행 번호는 2부터
                # G열(인덱스 6) = 진행상태
                if len(row) > SheetColumns.STATUS - 1:
                    status = row[SheetColumns.STATUS - 1]  # 0-based index
                    
                    # 완료가 아닌 첫 번째 행 반환
                    if status in processable_statuses:
                        return {
                            'row_index': idx,
                            'row_num': row[SheetColumns.ROW_NUM - 1] if len(row) >= SheetColumns.ROW_NUM else '',
                            'upper_path': row[SheetColumns.UPPER_PATH - 1] if len(row) >= SheetColumns.UPPER_PATH else '',
                            'sub_path': row[SheetColumns.SUB_PATH - 1] if len(row) >= SheetColumns.SUB_PATH else '',
                            'file_name': row[SheetColumns.FILE_NAME - 1] if len(row) >= SheetColumns.FILE_NAME else '',
                            'file_type': row[SheetColumns.FILE_TYPE - 1] if len(row) >= SheetColumns.FILE_TYPE else '',
                            'status': status,  # 현재 상태 추가
                        }
            
            return None  # 처리할 작업 없음
            
        except Exception as e:
            print(f"❌ 작업 조회 실패: {e}")
            return None

    def get_completed_tasks(self):
        """
        진행상태가 '완료'인 모든 작업을 반환합니다.
        
        Returns:
            list: 완료된 작업 리스트 (row_index, upper_path, sub_path, file_name, status 등)
        """
        try:
            all_values = self.sheet.get_all_values()
            completed_tasks = []
            
            for idx, row in enumerate(all_values[1:], start=2):
                if len(row) > SheetColumns.STATUS - 1:
                    status = row[SheetColumns.STATUS - 1]
                    if status == Status.COMPLETED:
                        completed_tasks.append({
                            'row_index': idx,
                            'row_num': row[SheetColumns.ROW_NUM - 1] if len(row) >= SheetColumns.ROW_NUM else '',
                            'upper_path': row[SheetColumns.UPPER_PATH - 1] if len(row) >= SheetColumns.UPPER_PATH else '',
                            'sub_path': row[SheetColumns.SUB_PATH - 1] if len(row) >= SheetColumns.SUB_PATH else '',
                            'file_name': row[SheetColumns.FILE_NAME - 1] if len(row) >= SheetColumns.FILE_NAME else '',
                            'file_type': row[SheetColumns.FILE_TYPE - 1] if len(row) >= SheetColumns.FILE_TYPE else '',
                            'status': status,
                        })
            
            return completed_tasks
            
        except Exception as e:
            print(f"❌ 완료 작업 조회 실패: {e}")
            return []
    
    def _api_call_with_retry(self, func, *args, **kwargs):
        """
        Rate Limit 대응을 위한 재시도 래퍼 함수
        
        Args:
            func: 실행할 함수
            *args, **kwargs: 함수 인자
            
        Returns:
            함수 실행 결과
            
        Raises:
            Exception: 최대 재시도 후에도 실패 시
        """
        last_exception = None
        
        for attempt in range(SHEETS_API_RETRY_COUNT):
            try:
                time.sleep(SHEETS_API_MIN_DELAY)  # 최소 대기
                return func(*args, **kwargs)
            except Exception as e:
                last_exception = e
                error_str = str(e)
                
                # Rate Limit (429) 에러인 경우 대기 후 재시도
                if '429' in error_str or 'Quota exceeded' in error_str:
                    wait_time = SHEETS_API_RETRY_DELAY * (attempt + 1)
                    print(f"   ⏳ API 한도 초과, {wait_time}초 대기 후 재시도 ({attempt + 1}/{SHEETS_API_RETRY_COUNT})...")
                    time.sleep(wait_time)
                else:
                    # 다른 종류의 에러는 바로 raise
                    raise
        
        # 모든 재시도 실패
        raise last_exception

    def update_status(self, row_index, status):
        """
        진행상태(G열)를 업데이트합니다.
        
        Rate Limit 대응을 위한 재시도 로직이 포함되어 있습니다.
        
        Args:
            row_index (int): 행 번호 (1-based)
            status (str): 상태값 (대기/진행중/완료/오류)
            
        Returns:
            bool: 성공 여부
        """
        try:
            self._api_call_with_retry(
                self.sheet.update_cell, row_index, SheetColumns.STATUS, status
            )
            return True
        except Exception as e:
            print(f"❌ 상태 업데이트 실패: {e}")
            return False
    
    def update_file_name(self, row_index, new_file_name):
        """
        파일명(D열)을 업데이트합니다.
        
        확장자가 대문자인 경우 소문자로 정규화할 때 사용합니다.
        
        Args:
            row_index (int): 행 번호 (1-based)
            new_file_name (str): 새 파일명 (확장자 소문자)
        """
        try:
            self._api_call_with_retry(
                self.sheet.update_cell, row_index, SheetColumns.FILE_NAME, new_file_name
            )
            print(f"   ✅ 시트 파일명 업데이트 완료")
        except Exception as e:
            print(f"⚠️ 파일명 업데이트 실패: {e}")
    
    def set_start_time(self, row_index):
        """
        시작일시(H열)에 현재 시간을 기록합니다.
        
        Args:
            row_index (int): 행 번호 (1-based)
        """
        try:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.sheet.update_cell(row_index, SheetColumns.START_TIME, now)
        except Exception as e:
            print(f"❌ 시작시간 기록 실패: {e}")
    
    def set_end_time(self, row_index):
        """
        종료일시(I열)에 현재 시간을 기록합니다.
        
        Args:
            row_index (int): 행 번호 (1-based)
        """
        try:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.sheet.update_cell(row_index, SheetColumns.END_TIME, now)
        except Exception as e:
            print(f"❌ 종료시간 기록 실패: {e}")
    
    # ==========================================================================
    # [진행 상황 추적] 토큰 사용량
    # ==========================================================================
    
    def get_current_tokens(self, row_index):
        """
        현재 토큰 사용량을 가져옵니다 (L, M열).
        
        Args:
            row_index (int): 행 번호 (1-based)
            
        Returns:
            dict: {'input_tokens': int, 'output_tokens': int}
        """
        try:
            row_data = self.sheet.row_values(row_index)
            
            def safe_int(value):
                """문자열을 안전하게 정수로 변환"""
                if not value:
                    return 0
                try:
                    return int(str(value).replace(',', '').strip())
                except (ValueError, TypeError):
                    return 0
            
            input_tokens = safe_int(row_data[SheetColumns.INPUT_TOKEN - 1]) if len(row_data) >= SheetColumns.INPUT_TOKEN else 0
            output_tokens = safe_int(row_data[SheetColumns.OUTPUT_TOKEN - 1]) if len(row_data) >= SheetColumns.OUTPUT_TOKEN else 0
            
            return {
                'input_tokens': input_tokens,
                'output_tokens': output_tokens
            }
            
        except Exception as e:
            print(f"⚠️ 토큰 사용량 조회 실패: {e}")
            return {'input_tokens': 0, 'output_tokens': 0}
    
    def update_tokens(self, row_index, input_tokens, output_tokens):
        """
        토큰 사용량을 업데이트합니다 (L, M열).
        기존 값에 누적합니다.
        
        Args:
            row_index (int): 행 번호 (1-based)
            input_tokens (int): 이번에 사용한 인풋 토큰 수
            output_tokens (int): 이번에 사용한 아웃풋 토큰 수
        """
        try:
            current = self.get_current_tokens(row_index)
            
            new_input_tokens = current['input_tokens'] + input_tokens
            new_output_tokens = current['output_tokens'] + output_tokens
            
            self.sheet.update_cell(row_index, SheetColumns.INPUT_TOKEN, new_input_tokens)
            self.sheet.update_cell(row_index, SheetColumns.OUTPUT_TOKEN, new_output_tokens)
            
        except Exception as e:
            print(f"⚠️ 토큰 업데이트 실패: {e}")
    
    def reset_tokens(self, row_index):
        """
        토큰 사용량을 초기화합니다 (L, M열).
        새 작업 시작 시 호출합니다.
        
        Args:
            row_index (int): 행 번호 (1-based)
        """
        try:
            self.sheet.update_cell(row_index, SheetColumns.INPUT_TOKEN, 0)
            self.sheet.update_cell(row_index, SheetColumns.OUTPUT_TOKEN, 0)
        except Exception as e:
            print(f"⚠️ 토큰 초기화 실패: {e}")
    
    # ==========================================================================
    # [오류 및 완료 처리]
    # ==========================================================================
    
    def record_error(self, row_index, error_message, module_name="Unknown"):
        """
        오류 정보를 기록합니다.
        - G열: 상태를 '오류'로 변경
        - L열: 오류 상세 내용 기록
        
        Args:
            row_index (int): 행 번호 (1-based)
            error_message (str): 오류 메시지
            module_name (str): 오류 발생 모듈명
        """
        try:
            # G열: 상태를 '오류'로 변경
            self.update_status(row_index, Status.ERROR)
            
            # L열: 오류 상세 기록
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            error_detail = f"[{now}] 모듈: {module_name}\n오류: {error_message}"
            
            self.sheet.update_cell(row_index, SheetColumns.ERROR, error_detail)
            
        except Exception as e:
            print(f"❌ 오류 기록 실패: {e}")
    
    def mark_completed(self, row_index):
        """
        작업 완료 처리를 합니다.
        - G열: 상태를 '완료'로 변경
        - I열: 종료시간 기록
        
        Args:
            row_index (int): 행 번호 (1-based)
        """
        try:
            self.update_status(row_index, Status.COMPLETED)
            self.set_end_time(row_index)
        except Exception as e:
            print(f"❌ 완료 처리 실패: {e}")
    
    def start_task(self, row_index):
        """
        작업 시작 처리를 합니다.
        - G열: 상태를 '진행중'으로 변경
        - H열: 시작시간 기록
        - L, M열: 토큰 사용량 초기화
        
        Args:
            row_index (int): 행 번호 (1-based)
        """
        try:
            self.update_status(row_index, Status.IN_PROGRESS)
            self.set_start_time(row_index)
            self.reset_tokens(row_index)
        except Exception as e:
            print(f"❌ 작업 시작 처리 실패: {e}")
    
    # ==========================================================================
    # [통계 및 조회]
    # ==========================================================================
    
    def get_overall_progress(self):
        """
        전체 진행율을 계산합니다.
        
        Returns:
            float: 진행율 (0~100)
        """
        try:
            all_values = self.sheet.get_all_values()
            
            total_count = 0
            completed_count = 0
            
            # 헤더 제외 (1행은 헤더)
            for row in all_values[1:]:
                if len(row) > SheetColumns.STATUS - 1:
                    status = row[SheetColumns.STATUS - 1]
                    if status:  # 빈 행 제외
                        total_count += 1
                        if status == Status.COMPLETED:
                            completed_count += 1
            
            if total_count == 0:
                return 0.0
            
            return (completed_count / total_count) * 100
            
        except Exception as e:
            print(f"⚠️ 진행율 계산 실패: {e}")
            return 0.0
    
    def get_review_progress(self):
        """
        1차 검수완료 진행율을 계산합니다.
        
        전체 파일 중 "1차 검수완료" 상태인 파일의 비율을 반환합니다.
        
        Returns:
            float: 1차 검수완료 진행율 (0~100)
        """
        try:
            all_values = self.sheet.get_all_values()
            
            total_count = 0
            review_completed_count = 0
            
            # 헤더 제외 (1행은 헤더)
            for row in all_values[1:]:
                if len(row) > SheetColumns.STATUS - 1:
                    status = row[SheetColumns.STATUS - 1]
                    if status:  # 빈 행 제외
                        total_count += 1
                        if status == Status.REVIEW_1_COMPLETED:
                            review_completed_count += 1
            
            if total_count == 0:
                return 0.0
            
            return (review_completed_count / total_count) * 100
            
        except Exception as e:
            print(f"⚠️ 검수 진행율 계산 실패: {e}")
            return 0.0
    
    def get_task_times(self, row_index):
        """
        작업의 시작/종료 시간을 가져옵니다.
        
        Args:
            row_index (int): 행 번호 (1-based)
            
        Returns:
            dict: {'start_time': str, 'end_time': str}
        """
        try:
            row_data = self.sheet.row_values(row_index)
            
            start_time = row_data[SheetColumns.START_TIME - 1] if len(row_data) >= SheetColumns.START_TIME else ''
            end_time = row_data[SheetColumns.END_TIME - 1] if len(row_data) >= SheetColumns.END_TIME else ''
            
            return {
                'start_time': start_time,
                'end_time': end_time
            }
            
        except Exception as e:
            print(f"⚠️ 작업 시간 조회 실패: {e}")
            return {'start_time': '', 'end_time': ''}
