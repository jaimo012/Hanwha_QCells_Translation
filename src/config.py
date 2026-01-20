"""
설정 모듈 (Configuration)

프로젝트 전역에서 사용되는 설정값들을 관리합니다.
- API 키, 폴더 경로, 모델명, 배치 사이즈 등의 상수를 정의합니다.
"""

import os

# ==============================================================================
# [API 설정] ⭐ 환경변수 또는 아래에 직접 입력
# ==============================================================================
# 환경변수 GEMINI_API_KEY가 있으면 사용, 없으면 아래 값 사용
API_KEY = os.environ.get("GEMINI_API_KEY", "여기에_API_키를_입력하세요")

# ==============================================================================
# [모델 설정]
# ==============================================================================
MODEL_NAME = "gemini-2.5-flash"  # 또는 "gemini-1.5-flash"

# ==============================================================================
# [Google Sheets 설정] ⭐ 작업 관리 시트
# ==============================================================================
GOOGLE_SHEETS_URL = "https://docs.google.com/spreadsheets/d/1xYby26nGoyXC3tGk1b3BqSNMl3QssNCqDFkCXAnRIs0/edit?gid=0#gid=0"
GOOGLE_SHEETS_NAME = "RAW"  # 시트 이름

# ==============================================================================
# [폴더 경로 설정]
# ==============================================================================
# 프로젝트 루트 디렉토리 (src의 상위 폴더)
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 데이터 폴더 (data 폴더 하위)
DATA_FOLDER = os.path.join(PROJECT_ROOT, "data")

# 원본 파일 폴더: 번역할 원본 파일이 있는 곳
ORIGIN_FOLDER = os.path.join(DATA_FOLDER, "origin_data_folder")

# 완료 폴더: 번역이 완료된 파일이 저장되는 곳
COMPLETED_FOLDER = os.path.join(DATA_FOLDER, "completed_folder")

# (기존 호환용) INPUT_FOLDER는 ORIGIN_FOLDER와 동일
INPUT_FOLDER = ORIGIN_FOLDER

# ==============================================================================
# [번역 설정]
# ==============================================================================
# 한 번에 번역할 텍스트 개수 (API 호출 효율성과 정확도의 균형)
# 파일 형식별로 다르게 설정 가능
BATCH_SIZE_DOCX = 30   # Word: 문단이 길어서 작게
BATCH_SIZE_PPTX = 70   # PowerPoint: 짧은 텍스트가 많아서 크게
BATCH_SIZE_XLSX = 70   # Excel: 셀 단위라 크게

# 기본 배치 사이즈 (하위 호환용)
BATCH_SIZE = 30

# API 호출 간 대기 시간 (초) - Rate Limit 방지
API_DELAY_SECONDS = 0.5

# 중간 저장 주기 (배치 횟수 기준)
AUTO_SAVE_INTERVAL = 10

# ==============================================================================
# [슬랙 웹훅]
# ==============================================================================
# 환경변수 SLACK_WEBHOOK_URL이 있으면 사용, 없으면 빈 문자열 (알림 비활성화)
slack_webhooks = os.environ.get("SLACK_WEBHOOK_URL", "")

# ==============================================================================
# [용어집 설정]
# ==============================================================================
# 용어정의 파일 경로
GLOSSARY_FILE_NAME = "용어정의.xlsx"

# 프롬프트에 포함할 최대 용어 수 (토큰 비용 관리)
GLOSSARY_MAX_TERMS = 200

# ==============================================================================
# [지원 파일 형식]
# ==============================================================================
# 번역 가능한 최신 형식
SUPPORTED_EXTENSIONS = ('.docx', '.xlsx', '.pptx')

# 변환 후 번역 가능한 구버전 형식
CONVERTIBLE_EXTENSIONS = ('.doc',)  # .doc → .docx 변환 후 번역

# 전체 지원 형식 (최신 + 구버전)
ALL_SUPPORTED_EXTENSIONS = SUPPORTED_EXTENSIONS + CONVERTIBLE_EXTENSIONS

# ==============================================================================
# [초기화 함수]
# ==============================================================================
def ensure_folders_exist():
    """
    필요한 폴더들이 존재하는지 확인하고, 없으면 생성합니다.
    """
    os.makedirs(DATA_FOLDER, exist_ok=True)
    os.makedirs(ORIGIN_FOLDER, exist_ok=True)
    os.makedirs(COMPLETED_FOLDER, exist_ok=True)


def validate_config():
    """
    설정값들이 올바른지 검증합니다.
    
    Returns:
        tuple: (성공 여부, 오류 메시지)
    """
    if not API_KEY or API_KEY == "여기에_API_키를_입력하세요":
        return False, "API_KEY가 설정되지 않았습니다. src/config.py 파일에서 API_KEY를 입력해주세요."
    
    # 서비스 계정 파일 확인
    credentials_path = os.path.join(PROJECT_ROOT, "credentials.json")
    if not os.path.exists(credentials_path):
        return False, "credentials.json 파일이 없습니다. 프로젝트 루트에 서비스 계정 키 파일을 복사해주세요."
    
    return True, "설정이 올바릅니다."
