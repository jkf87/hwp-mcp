# -*- coding: utf-8 -*-

import os
import logging
from typing import Dict, List, Optional, Union, Any
import json

# pyhwpx import
try:
    from pyhwpx import Hwp
except ImportError:
    Hwp = None

logger = logging.getLogger("hwp-template-engine")

class HwpTemplateEngine:
    """
    pyhwpx를 사용하여 HWP 템플릿 작업을 수행하는 엔진 클래스
    """
    
    def __init__(self, visible: bool = True):
        """
        초기화
        
        Args:
            visible: HWP 창 표시 여부
        """
        if Hwp is None:
            raise ImportError("pyhwpx library is not installed.")
            
        self.hwp = Hwp(visible=visible)
        self.is_connected = True
        
    def open(self, path: str) -> bool:
        """
        HWP 파일 열기
        
        Args:
            path: 파일 경로
            
        Returns:
            성공 여부
        """
        try:
            if not os.path.exists(path):
                logger.error(f"File not found: {path}")
                return False
                
            self.hwp.open(path)
            return True
        except Exception as e:
            logger.error(f"Error opening file: {e}")
            return False
            
    def get_field_list(self) -> List[str]:
        """
        문서 내의 누름틀(필드) 목록을 가져옵니다.
        
        Returns:
            필드 이름 목록
        """
        try:
            # GetFieldList는 "field1" + chr(2) + "field2" 형태의 문자열을 반환
            field_str = self.hwp.GetFieldList(1, 2) # 1: 모든 필드, 2: 옵션
            if not field_str:
                return []
            
            # chr(2)로 구분된 문자열을 리스트로 변환
            fields = field_str.split(chr(2))
            # 중복 제거 및 정렬
            return sorted(list(set(fields)))
        except Exception as e:
            logger.error(f"Error getting field list: {e}")
            return []
            
    def inject_data(self, data: Dict[str, Any]) -> int:
        """
        필드에 데이터를 주입합니다.
        
        Args:
            data: 필드명-값 딕셔너리
            
        Returns:
            주입된 필드 개수
        """
        count = 0
        try:
            available_fields = self.get_field_list()
            
            for field, value in data.items():
                if field in available_fields:
                    # pyhwpx의 put_field_text 사용
                    # value가 리스트인 경우 등은 pyhwpx가 처리
                    self.hwp.put_field_text(field, value)
                    count += 1
                else:
                    logger.warning(f"Field not found: {field}")
                    
            return count
        except Exception as e:
            logger.error(f"Error injecting data: {e}")
            return count
            
    def save(self, path: str) -> bool:
        """
        파일 저장
        
        Args:
            path: 저장할 경로
            
        Returns:
            성공 여부
        """
        try:
            self.hwp.save_as(path)
            return True
        except Exception as e:
            logger.error(f"Error saving file: {e}")
            return False
            
    def quit(self):
        """HWP 종료"""
        try:
            self.hwp.quit()
            self.is_connected = False
        except Exception as e:
            logger.error(f"Error quitting: {e}")

    def process_template(self, template_path: str, data: Dict[str, Any], output_path: str) -> bool:
        """
        템플릿 처리 전체 프로세스 실행
        
        Args:
            template_path: 템플릿 파일 경로
            data: 주입할 데이터
            output_path: 결과 파일 저장 경로
            
        Returns:
            성공 여부
        """
        if not self.open(template_path):
            return False
            
        self.inject_data(data)
        
        return self.save(output_path)