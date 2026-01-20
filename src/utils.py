"""
유틸리티 모듈 (Utilities)

프로젝트 전반에서 사용되는 공통 유틸리티 함수들을 정의합니다.
"""

import re


def has_korean(text):
    """
    텍스트에 한글이 포함되어 있는지 확인합니다.
    
    Args:
        text: 확인할 텍스트
        
    Returns:
        bool: 한글이 포함되어 있으면 True, 아니면 False
        
    Examples:
        >>> has_korean("안녕하세요")
        True
        >>> has_korean("Hello World")
        False
        >>> has_korean("Hello 세계")
        True
    """
    if not isinstance(text, str):
        return False
    return re.search('[가-힣]', text) is not None

