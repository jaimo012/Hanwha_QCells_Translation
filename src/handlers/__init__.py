"""
핸들러 패키지 (Handlers Package)

각 파일 형식별 처리 모듈을 제공합니다.
- docx_handler: Word 문서 처리
- pptx_handler: PowerPoint 문서 처리
- xlsx_handler: Excel 문서 처리
"""

from .docx_handler import process_docx
from .pptx_handler import process_pptx
from .xlsx_handler import process_xlsx

__all__ = ['process_docx', 'process_pptx', 'process_xlsx']

