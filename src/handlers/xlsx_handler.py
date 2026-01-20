"""
Excel 문서 핸들러 (XLSX Handler)

Excel 문서(.xlsx)의 번역을 처리합니다.
- xlwings를 사용하여 고속 처리
- 수식 보존 및 안전 종료
- 토큰 사용량 추적
- 긴 경로 자동 처리 (임시 폴더 활용)
"""

import os
import time
import shutil
import tempfile
import xlwings as xw

from ..config import BATCH_SIZE_XLSX, AUTO_SAVE_INTERVAL
from ..utils import has_korean
from ..translator import translate_batch

# 저장 재시도 설정
SAVE_MAX_RETRIES = 3
SAVE_RETRY_DELAY = 2  # 초

# Excel 경로 길이 제한 (안전 마진 포함)
MAX_PATH_LENGTH = 180

# Excel에서 문제가 되는 특수문자
PROBLEMATIC_CHARS = ['[', ']', '<', '>', '?', '*', '|']

# Excel 셀 최대 문자 수 (32,767자이지만 안전 마진)
MAX_CELL_LENGTH = 32000


def get_short_temp_path(original_path):
    """
    문제가 있는 경로의 파일을 위한 짧은 임시 경로를 생성합니다.
    
    Args:
        original_path (str): 원본 파일 경로
        
    Returns:
        str: 짧은 임시 파일 경로
    """
    _, ext = os.path.splitext(original_path)
    # 임시 폴더에 짧은 이름으로 저장
    temp_dir = tempfile.gettempdir()
    temp_name = f"xltemp_{int(time.time())}{ext}"
    return os.path.join(temp_dir, temp_name)


def has_problematic_path(file_path):
    """
    경로에 Excel 저장 시 문제가 될 수 있는 요소가 있는지 확인합니다.
    
    Args:
        file_path (str): 파일 경로
        
    Returns:
        tuple: (문제 여부, 문제 사유)
    """
    # 경로 길이 체크
    if len(file_path) > MAX_PATH_LENGTH:
        return True, f"경로가 너무 깁니다 ({len(file_path)}자)"
    
    # 특수문자 체크
    for char in PROBLEMATIC_CHARS:
        if char in file_path:
            return True, f"경로에 특수문자 '{char}'가 포함되어 있습니다"
    
    return False, None


def save_workbook_simple(wb, file_path, max_retries=SAVE_MAX_RETRIES):
    """
    Excel 워크북을 단순 재시도 로직과 함께 저장합니다.
    (임시 경로 작업 중에 사용 - 이동 없이 그냥 저장)
    
    Args:
        wb: xlwings Workbook 객체
        file_path (str): 저장할 파일 경로
        max_retries (int): 최대 재시도 횟수
        
    Returns:
        bool: 저장 성공 여부
    """
    for attempt in range(max_retries):
        try:
            wb.save(file_path)
            return True
        except Exception as e:
            error_msg = str(e)
            print(f"\n      ⚠️ 저장 실패 (시도 {attempt + 1}/{max_retries}): {error_msg[:80]}")
            
            if attempt < max_retries - 1:
                print(f"      ⏳ {SAVE_RETRY_DELAY}초 후 재시도...")
                time.sleep(SAVE_RETRY_DELAY)
            else:
                print(f"      ❌ 최대 재시도 횟수 초과. 저장 실패.")
                return False
    
    return False


def move_file_with_retry(src_path, dst_path, max_retries=SAVE_MAX_RETRIES):
    """
    파일을 재시도 로직과 함께 이동합니다.
    
    Args:
        src_path (str): 원본 파일 경로
        dst_path (str): 대상 파일 경로
        max_retries (int): 최대 재시도 횟수
        
    Returns:
        bool: 이동 성공 여부
    """
    for attempt in range(max_retries):
        try:
            # 기존 파일이 있으면 삭제
            if os.path.exists(dst_path):
                os.remove(dst_path)
            shutil.move(src_path, dst_path)
            return True
        except Exception as e:
            error_msg = str(e)
            print(f"      ⚠️ 파일 이동 실패 (시도 {attempt + 1}/{max_retries}): {error_msg[:80]}")
            
            if attempt < max_retries - 1:
                print(f"      ⏳ {SAVE_RETRY_DELAY}초 후 재시도...")
                time.sleep(SAVE_RETRY_DELAY)
            else:
                print(f"      ❌ 파일 이동 최대 재시도 횟수 초과.")
                return False
    
    return False


def truncate_cell_value(value):
    """
    셀 값이 Excel 최대 길이를 초과하면 자릅니다.
    
    Args:
        value: 셀 값
        
    Returns:
        처리된 셀 값
    """
    if isinstance(value, str) and len(value) > MAX_CELL_LENGTH:
        return value[:MAX_CELL_LENGTH - 3] + "..."
    return value


def write_range_safely(used_range, all_values, sheet):
    """
    범위에 데이터를 안전하게 씁니다.
    전체 쓰기 실패 시 행 단위 쓰기로 폴백합니다.
    
    Args:
        used_range: xlwings Range 객체
        all_values: 2D 리스트 데이터
        sheet: xlwings Sheet 객체
        
    Returns:
        bool: 성공 여부
    """
    # 셀 값 길이 검증 및 자르기
    for row_idx, row in enumerate(all_values):
        if row is None:
            continue
        if not isinstance(row, list):
            continue
        for col_idx, val in enumerate(row):
            all_values[row_idx][col_idx] = truncate_cell_value(val)
    
    # 방법 1: 전체 범위 쓰기 시도
    try:
        used_range.value = all_values
        return True
    except Exception as e:
        print(f"      ⚠️ 전체 범위 쓰기 실패: {str(e)[:50]}")
        print(f"      🔄 행 단위 쓰기로 전환합니다...")
    
    # 방법 2: 행 단위 쓰기 (개별 셀보다 훨씬 빠름)
    try:
        error_count = 0
        total_rows = len(all_values)
        
        for row_idx, row in enumerate(all_values):
            if row is None:
                continue
            if not isinstance(row, list):
                row = [row]
            
            try:
                # 행 단위로 쓰기 (훨씬 빠름)
                col_count = len(row)
                row_range = sheet.range((row_idx + 1, 1), (row_idx + 1, col_count))
                row_range.value = row
            except Exception:
                error_count += 1
                if error_count <= 3:
                    print(f"      ⚠️ 행 {row_idx+1} 쓰기 실패")
            
            # 진행 상황 표시 (100행마다)
            if (row_idx + 1) % 100 == 0:
                print(f"      📝 행 단위 쓰기 진행: {row_idx + 1}/{total_rows}행", end="\r")
        
        if total_rows > 100:
            print()  # 줄바꿈
        
        if error_count > 0:
            print(f"      ⚠️ 총 {error_count}개 행 쓰기 실패 (무시하고 계속)")
        return True
        
    except Exception as e:
        print(f"      ❌ 행 단위 쓰기도 실패: {str(e)[:50]}")
        return False


def process_xlsx(file_path, context, sheets_manager=None, row_index=None):
    """
    Excel 문서를 번역합니다.
    
    xlwings를 사용하여 백그라운드에서 고속으로 처리합니다.
    수식이 포함된 셀은 자동으로 건너뛰어 데이터를 보호합니다.
    
    경로에 특수문자가 있으면 임시 경로에서 작업 후 이동합니다.
    
    Args:
        file_path (str): 원본 Excel 파일 경로
        context (str): 번역 지침 (Context)
        sheets_manager (SheetsManager, optional): 시트 관리자 (토큰 추적용)
        row_index (int, optional): 시트 행 번호
        
    Returns:
        str: 번역된 파일의 경로 (성공 시)
        None: 실패 시
    """
    print(f"📗 Excel 처리 중: {os.path.basename(file_path)}")
    
    # 경로에 문제가 있는지 확인
    use_temp_path, reason = has_problematic_path(file_path)
    temp_work_path = None
    
    if use_temp_path:
        print(f"      ⚠️ {reason}")
        print(f"      📁 임시 경로에서 작업 후 완료 시 이동합니다...")
        temp_work_path = get_short_temp_path(file_path)
        # 원본 파일을 임시 경로로 복사
        shutil.copy2(file_path, temp_work_path)
        work_path = temp_work_path
    else:
        work_path = file_path
    
    # 최종 저장될 경로 (원래 경로)
    final_path = file_path
    
    # 앱 인스턴스 (백그라운드 실행)
    app = xw.App(visible=False)
    
    # [속도 최적화 핵심] 화면 갱신, 경고창, 수식 계산 끄기
    app.screen_updating = False
    app.display_alerts = False
    app.calculation = 'manual'

    try:
        wb = app.books.open(work_path)
        
        batch_cycle = 0
        total_translated_cells = 0
        total_input_tokens = 0
        total_output_tokens = 0
        
        for sheet_idx, sheet in enumerate(wb.sheets):
            print(f"\n   📊 시트 {sheet_idx + 1}/{len(wb.sheets)}: '{sheet.name}'")
            
            # [핵심] 안전하게 데이터 범위 가져오기
            all_values = None
            all_formulas = None
            used_range = None
            
            try:
                # 방법 1: used_range 먼저 시도 (가장 정확)
                used_range = sheet.used_range
                row_count = used_range.rows.count
                col_count = used_range.columns.count
                
                # 범위가 너무 크면 (10만 행 이상) 실제 데이터 범위 탐색
                if row_count > 100000:
                    print(f"      ⚠️ 시트 범위가 너무 큼 ({row_count}행). 실제 범위 탐색...")
                    
                    # 여러 열에서 마지막 행 찾기 (A, B, C, D, E열 중 최대값)
                    max_row = 1
                    for col_num in [1, 2, 3, 4, 5]:
                        try:
                            found_row = sheet.cells(1048576, col_num).end('up').row
                            if found_row > max_row:
                                max_row = found_row
                        except:
                            pass
                    
                    row_count = min(max_row, 50000)  # 최대 5만 행
                    col_count = min(col_count, 100)  # 최대 100열
                    
                    print(f"      ✅ 실제 범위로 조정: {row_count}행 x {col_count}열")
                    used_range = sheet.range((1, 1), (row_count, col_count))
                
                # 범위 제한 (열이 너무 많은 경우)
                elif col_count > 100:
                    col_count = 100
                    used_range = sheet.range((1, 1), (row_count, col_count))
                
                all_values = used_range.value
                all_formulas = used_range.formula
                    
            except Exception as e:
                error_msg = str(e)
                print(f"      ❌ 시트 데이터 로드 실패: {error_msg[:60]}")
                
                # 메모리 오류인 경우 작은 범위로 재시도
                if "메모리" in error_msg or "memory" in error_msg.lower():
                    print(f"      🔄 메모리 오류 - 작은 범위(5000행)로 재시도...")
                    try:
                        used_range = sheet.range('A1:AZ5000')
                        all_values = used_range.value
                        all_formulas = used_range.formula
                        print(f"      ✅ 작은 범위 로드 성공")
                    except Exception as e2:
                        print(f"      ❌ 재시도도 실패. 이 시트를 건너뜁니다: {str(e2)[:40]}")
                        continue
                else:
                    print(f"      ⚠️ 이 시트를 건너뜁니다.")
                    continue
            
            # 데이터가 없으면 건너뛰기
            if all_values is None:
                print(f"      ⚠️ 데이터 없음. 건너뜀.")
                continue
            
            # [핵심] 2D 리스트로 안전하게 변환
            # Case 1: 단일 값 (셀 1개)
            if not isinstance(all_values, list):
                all_values = [[all_values]]
                all_formulas = [[all_formulas]] if all_formulas is not None else [[None]]
            # Case 2: 1행 데이터 (1D 리스트)
            elif all_values and not isinstance(all_values[0], list):
                all_values = [all_values]
                all_formulas = [all_formulas] if all_formulas is not None else [None]
            
            # all_formulas가 None인 경우 빈 2D 리스트로
            if all_formulas is None:
                all_formulas = [[None] * len(row) for row in all_values]
            
            # 빈 시트 건너뛰기
            if not all_values or len(all_values) == 0:
                print(f"      ⚠️ 빈 시트. 건너뜀.")
                continue
            
            # 첫 번째 행이 None이거나 빈 경우 체크
            if all_values[0] is None:
                print(f"      ⚠️ 데이터 형식 오류. 건너뜀.")
                continue
            
            # 열 수 계산 (첫 행 기준)
            first_row = all_values[0]
            col_count = len(first_row) if isinstance(first_row, list) else 1
            
            print(f"      📋 처리 범위: {len(all_values)}행 x {col_count}열")
            
            # 번역 대상 수집 (좌표와 텍스트)
            batch_coords = []  # (row, col) 좌표
            batch_texts = []
            
            for row_idx, row in enumerate(all_values):
                # row가 None이거나 리스트가 아닌 경우 건너뛰기
                if row is None:
                    continue
                if not isinstance(row, list):
                    row = [row]  # 단일 값을 리스트로 변환
                
                for col_idx, val in enumerate(row):
                    # 수식 체크 (안전한 인덱스 접근)
                    formula = None
                    try:
                        if all_formulas and row_idx < len(all_formulas):
                            formula_row = all_formulas[row_idx]
                            if formula_row is not None:
                                if isinstance(formula_row, list) and col_idx < len(formula_row):
                                    formula = formula_row[col_idx]
                                elif not isinstance(formula_row, list):
                                    formula = formula_row if col_idx == 0 else None
                    except (IndexError, TypeError):
                        formula = None
                    
                    if formula and isinstance(formula, str) and formula.startswith('='):
                        continue
                    
                    # 한글 문자열만 수집
                    if val and isinstance(val, str) and has_korean(val):
                        batch_coords.append((row_idx, col_idx))
                        batch_texts.append(val)
                        
                        # 배치 번역 실행
                        if len(batch_texts) >= BATCH_SIZE_XLSX:
                            translated, input_tokens, output_tokens = translate_batch(batch_texts, context)
                            
                            total_input_tokens += input_tokens
                            total_output_tokens += output_tokens
                            
                            if len(translated) == len(batch_coords):
                                # 번역 결과를 메모리에 반영
                                for (r, c), txt in zip(batch_coords, translated):
                                    all_values[r][c] = txt
                            
                            total_translated_cells += len(translated)
                            batch_cycle += 1
                            
                            print(f"   ▶ 배치 {batch_cycle}회 진행 중...          ", end="\r")
                            
                            if batch_cycle % AUTO_SAVE_INTERVAL == 0:
                                print()
                                print(f"   💾 [자동저장] 엑셀 중간 저장...")
                                # 안전한 범위 쓰기 (실패 시 개별 셀 쓰기로 폴백)
                                if not write_range_safely(used_range, all_values, sheet):
                                    print(f"      ⚠️ 데이터 쓰기 실패, 저장 건너뜀")
                                    continue
                                # 중간 저장은 작업 경로(임시 또는 원본)에 직접 저장
                                if not save_workbook_simple(wb, work_path):
                                    print(f"      ⚠️ 중간 저장 실패, 계속 진행...")
                                else:
                                    print(f"      ✅ 중간 저장 완료")
                                
                                if sheets_manager and row_index:
                                    sheets_manager.update_tokens(row_index, total_input_tokens, total_output_tokens)
                                    total_input_tokens = 0
                                    total_output_tokens = 0
                            
                            batch_coords = []
                            batch_texts = []
                            time.sleep(0.2)
            
            # 잔여 데이터 처리
            if batch_texts:
                print(f"\n   🔄 [Sheet: {sheet.name}] 잔여 {len(batch_texts)}개 처리 중...")
                translated, input_tokens, output_tokens = translate_batch(batch_texts, context)
                
                total_input_tokens += input_tokens
                total_output_tokens += output_tokens
                
                if len(translated) == len(batch_coords):
                    for (r, c), txt in zip(batch_coords, translated):
                        all_values[r][c] = txt
                    total_translated_cells += len(translated)
                    batch_cycle += 1
                print(f"   ✅ [Sheet: {sheet.name}] 잔여 처리 완료")
            
            # 시트 데이터 한 번에 쓰기 (안전 모드)
            if not write_range_safely(used_range, all_values, sheet):
                print(f"      ⚠️ 시트 '{sheet.name}' 데이터 쓰기 실패")
        
        # 모든 시트 처리 완료 후 최종 저장
        print(f"\n   💾 최종 저장 중...")
        if not save_workbook_simple(wb, work_path):
            raise Exception("최종 저장 실패 - 재시도 횟수 초과")
        
        # 워크북 닫기 (파일 잠금 해제)
        wb.close()
        print(f"   ✅ 워크북 저장 완료")
        
        # 임시 경로 사용한 경우, 원래 위치로 이동
        if use_temp_path and temp_work_path:
            print(f"   📦 파일을 원래 위치로 이동 중...")
            if not move_file_with_retry(temp_work_path, final_path):
                raise Exception("파일 이동 실패")
            print(f"   ✅ 파일 이동 완료")
        
        # 최종 토큰 사용량 업데이트
        if sheets_manager and row_index:
            if total_input_tokens > 0 or total_output_tokens > 0:
                sheets_manager.update_tokens(
                    row_index,
                    total_input_tokens,
                    total_output_tokens
                )
        
        print()  # 진행 상황 줄 종료
        print(f"   ✅ Excel 번역 완료: {batch_cycle}개 배치, {total_translated_cells}개 셀")
        
    except Exception as e:
        error_msg = str(e)
        print(f"\n   ❌ Excel Error: {error_msg}")
        
        # 오류 유형별 추가 안내
        if "액세스" in error_msg or "access" in error_msg.lower():
            print("   💡 힌트: 파일이 다른 프로그램에서 열려 있을 수 있습니다.")
            print("   💡 Excel을 모두 닫고 다시 시도해주세요.")
        elif "218" in error_msg:
            print("   💡 힌트: 파일 경로가 218자를 초과합니다.")
        elif "RPC" in error_msg:
            print("   💡 힌트: Excel 연결이 끊어졌습니다. 프로그램을 다시 시작해주세요.")
        
        try:
            wb.close()
        except:
            pass
        
        # 임시 파일 정리
        if temp_work_path and os.path.exists(temp_work_path):
            try:
                os.remove(temp_work_path)
            except:
                pass
        
        raise Exception(f"번역 처리 실패: {error_msg[:50]}")
        
    finally:
        # [안전 종료] 설정 복구 및 프로세스 종료
        try:
            app.calculation = 'automatic'
            app.screen_updating = True
            app.display_alerts = True
            app.quit()
        except:
            # app이 이미 종료된 경우 무시
            pass
    
    return final_path
