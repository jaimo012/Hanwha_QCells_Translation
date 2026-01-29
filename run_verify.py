"""
검수 실행 스크립트 (Verify Runner Script)

1차 번역이 완료된 파일들의 검수를 실행합니다.
완료 상태인 파일에서 남아있는 한글을 추가 번역합니다.

사용법:
    python run_verify.py
"""

from src.verify import main

if __name__ == "__main__":
    main()
