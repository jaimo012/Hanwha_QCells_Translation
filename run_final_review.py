"""
최종검수 실행 스크립트 (Final Review Runner Script)

완료된 번역 파일들의 최종 검수를 실행합니다.
- 원본/번역본 파일 존재 여부 확인
- 파일 오픈 가능 여부 확인
- 번역 완료 여부 (한글 잔존 확인)
- Google Sheets '최종검수' 시트에 결과 기록

사용법:
    python run_final_review.py
"""

from src.final_review import main

if __name__ == "__main__":
    main()
