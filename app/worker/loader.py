import csv, json, os
from typing import List, Dict

class FileLoader:
    """파일을 로딩합니다."""

    @classmethod
    def load(cls, path: str) -> List[Dict[str, str]]:
        """경로 기반으로 데이터 추출"""
        _, ext = os.path.splitext(path)
        ext = ext.lower()

        if ext == ".csv":
            return cls._load_csv(path)
        elif ext == ".json":
            return cls._load_json(path)
        else:
            raise ValueError(f"지원하지 않는 파일 형식: {ext}")
        
    @staticmethod
    def _load_csv(path: str) -> List[Dict[str, str]]:
        with open(path, mode="r", encoding="utf-8-sig") as f:
            return [{k: v.strip() for k, v in row.items()} for row in csv.DictReader(f)]
        
    @staticmethod
    def _load_json(path: str) -> List[Dict[str, str]]:
        with open(path, mode="r", encoding="utf-8") as f:
            return json.load(f)