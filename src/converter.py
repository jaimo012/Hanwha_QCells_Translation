"""
파일 변환 모듈 (Converter)

구버전 Office 파일을 최신 형식으로 변환합니다.
- .doc → .docx (Word)
- .ppt → .pptx (PowerPoint) [필요시 추가 가능]
- .xls → .xlsx (Excel) [필요시 추가 가능]
"""

import os
import win32com.client as win32


def convert_doc_to_docx(doc_path, docx_path=None):
    """
    .doc 파일을 .docx 파일로 변환합니다.
    
    Microsoft Word를 사용하여 변환하므로 Word가 설치되어 있어야 합니다.
    
    Args:
        doc_path (str): 원본 .doc 파일 경로
        docx_path (str, optional): 저장할 .docx 파일 경로. 
                                   None이면 같은 위치에 확장자만 변경
        
    Returns:
        str: 변환된 .docx 파일 경로 (성공 시)
        None: 실패 시
        
    Raises:
        FileNotFoundError: 원본 파일이 없을 때
        Exception: Word 변환 실패 시
    """
    # 원본 파일 존재 확인
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {doc_path}")
    
    # 출력 경로 설정
    if docx_path is None:
        docx_path = os.path.splitext(doc_path)[0] + ".docx"
    
    # 절대 경로로 변환 (COM 객체는 절대 경로 필요)
    doc_path = os.path.abspath(doc_path)
    docx_path = os.path.abspath(docx_path)
    
    word = None
    doc = None
    
    try:
        print(f"   🔄 .doc → .docx 변환 중...")
        
        # Word 애플리케이션 실행 (백그라운드)
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        
        # 문서 열기
        doc = word.Documents.Open(doc_path)
        
        # .docx 형식으로 저장 (FileFormat=16은 docx)
        doc.SaveAs2(docx_path, FileFormat=16)
        
        print(f"   ✅ 변환 완료: {os.path.basename(docx_path)}")
        
        return docx_path
        
    except Exception as e:
        print(f"   ❌ .doc 변환 실패: {e}")
        raise
        
    finally:
        # 리소스 정리 (필수!)
        # COM 객체는 __len__ 메서드가 없어서 `if doc:` 대신 `is not None` 사용
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass  # 이미 닫힌 경우 무시
        
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass  # 이미 종료된 경우 무시


def needs_conversion(file_path):
    """
    파일이 변환이 필요한 구버전 형식인지 확인합니다.
    
    Args:
        file_path (str): 파일 경로
        
    Returns:
        bool: 변환 필요 여부
    """
    ext = os.path.splitext(file_path)[1].lower()
    # 변환이 필요한 구버전 확장자 목록
    old_formats = ['.doc']  # 필요시 '.ppt', '.xls' 추가
    return ext in old_formats


def get_converted_extension(file_path):
    """
    구버전 파일의 새 확장자를 반환합니다.
    
    Args:
        file_path (str): 파일 경로
        
    Returns:
        str: 새 확장자 (예: '.docx')
        None: 변환 대상이 아닐 때
    """
    ext = os.path.splitext(file_path)[1].lower()
    conversion_map = {
        '.doc': '.docx',
        # '.ppt': '.pptx',  # 필요시 추가
        # '.xls': '.xlsx',  # 필요시 추가
    }
    return conversion_map.get(ext)

