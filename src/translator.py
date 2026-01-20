"""
번역 모듈 (Translator)

Google Gemini API를 사용하여 번역을 수행하는 함수들을 정의합니다.
- Context 생성: 문서의 맥락을 분석
- 배치 번역: 텍스트 리스트를 일괄 번역
- 토큰 사용량 추적
- 용어집(Glossary) 적용
- 타임아웃 및 재시도 로직
"""

import json
import time
import google.generativeai as genai

from .config import API_KEY, MODEL_NAME, GLOSSARY_MAX_TERMS
from .prompts import PROMPT_CONTEXT_ANALYSIS, PROMPT_TRANSLATION_SYSTEM
from .glossary import get_glossary_prompt_text


# ==============================================================================
# API 설정
# ==============================================================================
genai.configure(api_key=API_KEY)

# 타임아웃 및 재시도 설정
API_TIMEOUT_SECONDS = 120  # 2분 타임아웃
MAX_RETRIES = 3            # 최대 재시도 횟수
RETRY_DELAY_SECONDS = 5    # 재시도 간 대기 시간

# 용어집 텍스트 (모듈 로드 시 한 번만 생성)
_glossary_text = None


def _get_glossary_text():
    """용어집 텍스트를 가져옵니다 (캐싱)."""
    global _glossary_text
    if _glossary_text is None:
        _glossary_text = get_glossary_prompt_text(max_terms=GLOSSARY_MAX_TERMS)
        if _glossary_text:
            print(f"   📚 용어집 프롬프트 로드 완료")
        else:
            print(f"   ⚠️ 용어집이 비어있거나 로드되지 않았습니다")
            _glossary_text = "(용어집 없음)"
    return _glossary_text


def generate_context(text_sample):
    """
    문서 샘플 텍스트를 분석하여 번역 지침(Context)을 생성합니다.
    
    Args:
        text_sample (str): 분석할 문서의 샘플 텍스트
        
    Returns:
        str: 생성된 번역 지침 (Context)
        
    Note:
        - 최대 10,000자까지만 분석합니다.
        - 오류 발생 시 기본 Context를 반환합니다.
        - 타임아웃 적용됨.
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            model = genai.GenerativeModel(MODEL_NAME)
            
            # 사용자 프롬프트에 텍스트 주입 (최대 10,000자)
            prompt = PROMPT_CONTEXT_ANALYSIS.format(
                extracted_text=text_sample[:10000]
            )
            
            response = model.generate_content(
                prompt,
                request_options={"timeout": API_TIMEOUT_SECONDS}
            )
            return response.text.strip()
            
        except Exception as e:
            error_msg = str(e)
            print(f"   ⚠️ Context 생성 오류 (시도 {attempt}/{MAX_RETRIES}): {error_msg[:100]}")
            
            if attempt < MAX_RETRIES:
                wait_time = RETRY_DELAY_SECONDS * attempt
                print(f"   🔄 {wait_time}초 후 재시도...")
                time.sleep(wait_time)
            else:
                print(f"   ❌ Context 생성 실패, 기본값 사용")
                return "MES Technical Document. Use standard terminology."
    
    return "MES Technical Document. Use standard terminology."


def translate_batch(text_list, file_context):
    """
    텍스트 리스트를 일괄 번역합니다.
    
    Args:
        text_list (list): 번역할 텍스트들의 리스트
        file_context (str): 번역 지침 (Context)
        
    Returns:
        tuple: (번역된 텍스트 리스트, 인풋 토큰 수, 아웃풋 토큰 수)
        
    Note:
        - 입력과 동일한 순서와 길이의 리스트를 반환합니다.
        - 오류 발생 시 원본 리스트와 토큰 0을 반환합니다.
        - 용어집(Glossary)이 프롬프트에 자동 포함됩니다.
        - 타임아웃 및 재시도 로직이 포함되어 있습니다.
    """
    if not text_list:
        return [], 0, 0
    
    # 용어집 텍스트 가져오기
    glossary_text = _get_glossary_text()
    
    # 리스트를 JSON 문자열로 변환
    json_input = json.dumps(text_list, ensure_ascii=False)
    
    # 프롬프트 생성
    prompt = PROMPT_TRANSLATION_SYSTEM.format(
        glossary_text=glossary_text,
        file_context=file_context,
        json_batch_list=json_input
    )
    
    # 재시도 로직
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            model = genai.GenerativeModel(MODEL_NAME)
            
            # 타임아웃 설정과 함께 API 호출
            response = model.generate_content(
                prompt,
                request_options={"timeout": API_TIMEOUT_SECONDS}
            )
            
            result_text = response.text.strip()
            
            # 토큰 사용량 추출
            input_tokens = 0
            output_tokens = 0
            
            if hasattr(response, 'usage_metadata'):
                usage = response.usage_metadata
                input_tokens = getattr(usage, 'prompt_token_count', 0) or 0
                output_tokens = getattr(usage, 'candidates_token_count', 0) or 0
            
            # 마크다운 제거 (안전장치)
            if result_text.startswith("```"):
                result_text = result_text.replace("```json", "").replace("```", "")
            
            translated_list = json.loads(result_text)
            
            return translated_list, input_tokens, output_tokens
            
        except json.JSONDecodeError as e:
            print(f"\n   ⚠️ JSON 파싱 오류 (시도 {attempt}/{MAX_RETRIES}): {e}")
            if attempt < MAX_RETRIES:
                print(f"   🔄 {RETRY_DELAY_SECONDS}초 후 재시도...")
                time.sleep(RETRY_DELAY_SECONDS)
            else:
                print(f"   ❌ 최대 재시도 횟수 초과, 원본 반환")
                return text_list, 0, 0
                
        except Exception as e:
            error_msg = str(e)
            
            # 타임아웃 또는 네트워크 오류 감지
            if "timeout" in error_msg.lower() or "deadline" in error_msg.lower():
                print(f"\n   ⏱️ API 타임아웃 (시도 {attempt}/{MAX_RETRIES})")
            elif "429" in error_msg or "quota" in error_msg.lower():
                print(f"\n   🚫 API 할당량 초과 (시도 {attempt}/{MAX_RETRIES})")
            else:
                print(f"\n   ❌ API 오류 (시도 {attempt}/{MAX_RETRIES}): {error_msg[:100]}")
            
            if attempt < MAX_RETRIES:
                wait_time = RETRY_DELAY_SECONDS * attempt  # 점진적 대기
                print(f"   🔄 {wait_time}초 후 재시도...")
                time.sleep(wait_time)
            else:
                print(f"   ❌ 최대 재시도 횟수 초과, 원본 반환")
                return text_list, 0, 0
    
    return text_list, 0, 0
