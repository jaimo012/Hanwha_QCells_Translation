"""
용어집 모듈 (Glossary)

용어정의 파일(xlsx)을 읽어서 번역 프롬프트에 주입할 형식으로 변환합니다.
- 한국어 → 영어 용어 매핑
- 프롬프트에 삽입할 텍스트 형식 생성
"""

import os
import pandas as pd

from .config import DATA_FOLDER


# 용어정의 파일 경로
GLOSSARY_FILE_PATH = os.path.join(DATA_FOLDER, "용어정의.xlsx")


class Glossary:
    """
    용어정의 파일을 관리하는 클래스입니다.
    
    Attributes:
        terms (list): 용어 목록 [(한국어, 영어, 설명), ...]
        is_loaded (bool): 용어집 로드 성공 여부
    """
    
    def __init__(self):
        self.terms = []
        self.is_loaded = False
        self._load_glossary()
    
    def _load_glossary(self):
        """
        용어정의 파일을 로드합니다.
        
        파일 구조 예상:
        - 1열: 한국어 용어
        - 2열: 영어 용어
        - 3열 이후: 설명/비고 (선택적)
        """
        try:
            if not os.path.exists(GLOSSARY_FILE_PATH):
                print(f"⚠️ 용어정의 파일이 없습니다: {GLOSSARY_FILE_PATH}")
                return
            
            # 엑셀 파일 읽기
            df = pd.read_excel(GLOSSARY_FILE_PATH)
            
            # 컬럼명 확인 및 정규화
            columns = df.columns.tolist()
            
            if len(columns) < 2:
                print("⚠️ 용어정의 파일에 최소 2개 이상의 컬럼이 필요합니다 (한국어, 영어)")
                return
            
            # 첫 번째와 두 번째 컬럼을 한국어/영어로 사용
            korean_col = columns[0]
            english_col = columns[1]
            
            # 용어 추출
            for _, row in df.iterrows():
                korean_term = str(row[korean_col]).strip() if pd.notna(row[korean_col]) else ""
                english_term = str(row[english_col]).strip() if pd.notna(row[english_col]) else ""
                
                # 빈 행 건너뛰기
                if not korean_term or not english_term:
                    continue
                
                # nan 문자열 제외
                if korean_term.lower() == 'nan' or english_term.lower() == 'nan':
                    continue
                
                self.terms.append((korean_term, english_term))
            
            self.is_loaded = True
            print(f"✅ 용어집 로드 완료: {len(self.terms)}개 용어")
            
        except Exception as e:
            print(f"❌ 용어집 로드 실패: {e}")
            self.is_loaded = False
    
    def get_prompt_text(self, max_terms=200):
        """
        프롬프트에 주입할 용어집 텍스트를 생성합니다.
        
        Args:
            max_terms (int): 프롬프트에 포함할 최대 용어 수 (토큰 비용 관리)
            
        Returns:
            str: 프롬프트용 용어집 텍스트
        """
        if not self.terms:
            return ""
        
        # 용어 수 제한 (토큰 비용 고려)
        terms_to_use = self.terms[:max_terms]
        
        # 테이블 형식으로 변환
        lines = ["| 한국어 | English |", "|--------|---------|"]
        
        for korean, english in terms_to_use:
            # 특수문자 이스케이프 (파이프 문자)
            korean_safe = korean.replace("|", "\\|")
            english_safe = english.replace("|", "\\|")
            lines.append(f"| {korean_safe} | {english_safe} |")
        
        return "\n".join(lines)
    
    def get_term_count(self):
        """용어 개수를 반환합니다."""
        return len(self.terms)
    
    def find_term(self, korean_text):
        """
        한국어 텍스트에서 용어집에 있는 용어를 찾습니다.
        
        Args:
            korean_text (str): 검색할 한국어 텍스트
            
        Returns:
            list: 발견된 용어 목록 [(한국어, 영어), ...]
        """
        found = []
        for korean, english in self.terms:
            if korean in korean_text:
                found.append((korean, english))
        return found


# 모듈 로드 시 싱글톤 인스턴스 생성
_glossary_instance = None


def get_glossary():
    """
    용어집 싱글톤 인스턴스를 반환합니다.
    
    Returns:
        Glossary: 용어집 인스턴스
    """
    global _glossary_instance
    if _glossary_instance is None:
        _glossary_instance = Glossary()
    return _glossary_instance


def get_glossary_prompt_text(max_terms=200):
    """
    프롬프트에 주입할 용어집 텍스트를 반환합니다.
    
    Args:
        max_terms (int): 최대 용어 수
        
    Returns:
        str: 프롬프트용 용어집 텍스트
    """
    glossary = get_glossary()
    return glossary.get_prompt_text(max_terms)

