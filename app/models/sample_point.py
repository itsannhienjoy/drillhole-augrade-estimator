from typing import Dict


class SamplePoint:
    hole_id: str
    x: float
    y: float
    length: float
    au_grade: float
    extra_data: Dict[str, str]
