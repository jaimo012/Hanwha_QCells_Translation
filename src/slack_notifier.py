"""
Slack 알림 모듈 (Slack Notifier)

Slack 웹훅을 통해 번역 완료/오류 알림을 전송합니다.
"""

import requests
from datetime import datetime

from .config import slack_webhooks


def send_slack_message(message):
    """
    Slack 웹훅으로 메시지를 전송합니다.
    
    Args:
        message (str): 전송할 메시지 (마크다운 형식 지원)
        
    Returns:
        bool: 성공 여부
    """
    # 웹훅 URL 확인
    if not slack_webhooks or slack_webhooks == '':
        print(f"   ⚠️ Slack 웹훅 URL이 설정되지 않았습니다.")
        return False
    
    try:
        payload = {"text": message}
        response = requests.post(slack_webhooks, json=payload, timeout=10)
        
        if response.status_code == 200:
            print(f"   📨 Slack 알림 전송 완료")
            return True
        else:
            print(f"   ⚠️ Slack 전송 실패 (HTTP {response.status_code}): {response.text}")
            return False
            
    except requests.exceptions.Timeout:
        print(f"   ⚠️ Slack 전송 타임아웃 (10초 초과)")
        return False
    except requests.exceptions.ConnectionError:
        print(f"   ⚠️ Slack 서버 연결 실패")
        return False
    except Exception as e:
        print(f"   ⚠️ Slack 전송 오류: {e}")
        return False


def format_datetime(dt_str):
    """
    날짜/시간 문자열을 포맷팅합니다.
    
    Args:
        dt_str (str): "yyyy-mm-dd HH:MM:SS" 형식의 문자열
        
    Returns:
        str: "yyyy.mm.dd HH:MM" 형식의 문자열
    """
    try:
        if not dt_str:
            return "-"
        dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%Y.%m.%d %H:%M")
    except:
        return dt_str


def calculate_duration(start_str, end_str):
    """
    시작/종료 시간으로 소요 시간을 계산합니다.
    
    Args:
        start_str (str): 시작 시간 문자열
        end_str (str): 종료 시간 문자열
        
    Returns:
        str: "00분 00초" 형식의 문자열
    """
    try:
        if not start_str or not end_str:
            return "-"
        
        start = datetime.strptime(start_str, "%Y-%m-%d %H:%M:%S")
        end = datetime.strptime(end_str, "%Y-%m-%d %H:%M:%S")
        
        duration = end - start
        total_seconds = int(duration.total_seconds())
        
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        
        if hours > 0:
            return f"{hours}시간 {minutes}분 {seconds}초"
        elif minutes > 0:
            return f"{minutes}분 {seconds}초"
        else:
            return f"{seconds}초"
            
    except:
        return "-"


def send_completion_notification(file_name, file_path, start_time, end_time, progress_percent):
    """
    번역 완료 알림을 Slack으로 전송합니다.
    
    Args:
        file_name (str): 파일명
        file_path (str): 파일 경로
        start_time (str): 시작 시간
        end_time (str): 종료 시간
        progress_percent (float): 전체 진행율 (0~100)
        
    Returns:
        bool: 전송 성공 여부
    """
    now = datetime.now().strftime("%Y.%m.%d %H:%M")
    duration = calculate_duration(start_time, end_time)
    start_formatted = format_datetime(start_time)
    end_formatted = format_datetime(end_time)
    
    message = f"""🔥 *한화큐셀 프로젝트 번역완료*
{now}

*전체 진행율*: {progress_percent:.1f}%

*파일명*
{file_name}

*경로*
{file_path}

*소요시간*: {duration} 소요
{start_formatted} ~ {end_formatted}

<https://docs.google.com/spreadsheets/d/1xYby26nGoyXC3tGk1b3BqSNMl3QssNCqDFkCXAnRIs0/edit?gid=0#gid=0|📂 *시트 바로가기*>
"""
    
    return send_slack_message(message)


def send_error_notification(error_message):
    """
    오류 발생 알림을 Slack으로 전송합니다.
    
    Args:
        error_message (str): 오류 내용
        
    Returns:
        bool: 전송 성공 여부
    """
    now = datetime.now().strftime("%Y.%m.%d %H:%M")
    
    # 오류 메시지가 너무 길면 자르기
    if len(error_message) > 500:
        error_message = error_message[:500] + "..."
    
    message = f"""🚨 *한화큐셀 프로젝트 번역 오류 발생!* <@U07C3D12E94>
{now}

{error_message}

<https://docs.google.com/spreadsheets/d/1xYby26nGoyXC3tGk1b3BqSNMl3QssNCqDFkCXAnRIs0/edit?gid=0#gid=0|📂 *시트 바로가기*>
"""
    
    return send_slack_message(message)


def send_review_completion_notification(file_name, file_path, review_progress_percent):
    """
    1차 검수 완료 알림을 Slack으로 전송합니다.
    
    Args:
        file_name (str): 파일명 ("-en" 포함)
        file_path (str): 파일 경로
        review_progress_percent (float): 1차 검수완료 진행율 (0~100)
        
    Returns:
        bool: 전송 성공 여부
    """
    now = datetime.now().strftime("%Y.%m.%d %H:%M")
    
    message = f"""🔧 *한화큐셀 프로젝트 1차 검수완료*
{now}

*1차 검수완료 진행율*: {review_progress_percent:.1f}%

*파일명*
{file_name}

*경로*
{file_path}

<https://docs.google.com/spreadsheets/d/1xYby26nGoyXC3tGk1b3BqSNMl3QssNCqDFkCXAnRIs0/edit?gid=0#gid=0|📂 *시트 바로가기*>
"""
    
    return send_slack_message(message)

