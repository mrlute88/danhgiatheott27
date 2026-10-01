"""
classifier.py - Agent 1: Data Classifier (Rule & Context Builder)
Standardizes student scores/completion levels according to Circular 27/2020/TT-BGDĐT
and builds pedagogical context and guidance for Agent 2.
"""

import re
from typing import Optional, List, Dict, Any
from pydantic import BaseModel, Field
from utils.tt27_knowledge import SUBJECT_COMPETENCIES


class StudentInput(BaseModel):
    student_name: str = Field(..., description="Full name of student")
    subject: str = Field("Toán", description="Subject name")
    grade: str = Field("Lớp 1", description="Grade level (Lớp 1 to Lớp 5)")
    score_or_level: Any = Field(..., description="Raw score or completion level from Excel (XL GK1, XL CK1, etc.)")
    exam_score: Optional[Any] = Field(default=None, description="Optional raw numerical exam score (KT CK1, KT GK1, etc.)")
    teacher_note: Optional[str] = Field(default="", description="Teacher qualitative notes or remarks")
    period: str = Field("Giữa Học kỳ 1", description="Evaluation period")
    tone: str = Field("Chuẩn mực & Động viên", description="Comment tone style")


class ClassificationResult(BaseModel):
    standardized_level: str = Field(..., description="T (Hoàn thành tốt), H (Hoàn thành), C (Chưa hoàn thành)")
    level_label: str = Field(..., description="Human readable Vietnamese level label")
    numerical_score: Optional[float] = Field(default=None, description="Parsed numerical score if available")
    tier: str = Field(..., description="Fine-grained performance tier")
    subject_focus_strengths: List[str] = Field(default_factory=list)
    subject_focus_growth: List[str] = Field(default_factory=list)
    guidance_notes: str = Field(..., description="Specific instructions for commentator agent")


class DataClassifierAgent:
    """Agent 1: Classifies student achievement level and generates instructional context."""

    @staticmethod
    def parse_score_and_level(raw_val: Any) -> tuple[str, Optional[float], str]:
        """
        Parses raw Excel cell value into (standardized_level, numerical_score, tier).
        Levels: 'T' (Tốt), 'H' (Hoàn thành), 'C' (Chưa hoàn thành).
        """
        if raw_val is None:
            return "H", None, "H_MID"

        val_str = str(raw_val).strip()
        if not val_str:
            return "H", None, "H_MID"

        # Check for numeric pattern (supports both '.' and ',' e.g., '9.5', '9,5')
        clean_num_str = val_str.replace(",", ".")
        match = re.search(r"(\d+(\.\d+)?)", clean_num_str)
        if match:
            try:
                num = float(match.group(1))
                if 0.0 <= num <= 10.0:
                    if num >= 9.0:
                        return "T", num, "T_HIGH" if num >= 9.5 else "T_NORMAL"
                    elif num >= 7.0:
                        return "H", num, "H_HIGH"
                    elif num >= 5.0:
                        return "H", num, "H_MID"
                    else:
                        return "C", num, "C_LOW"
            except ValueError:
                pass

        # String matching for completion levels
        val_lower = val_str.lower()
        if any(k in val_lower for k in ["htt", "tốt", "xuất sắc", "giỏi"]) or val_str.upper() == "T":
            return "T", None, "T_NORMAL"
        elif any(k in val_lower for k in ["cht", "chưa hoàn thành", "chưa đạt", "yếu", "kém"]) or val_str.upper() == "C":
            return "C", None, "C_LOW"
        elif any(k in val_lower for k in ["ht", "hoàn thành", "đạt", "khá", "trung bình"]) or val_str.upper() == "H":
            return "H", None, "H_MID"

        # Default fallback
        return "H", None, "H_MID"

    def classify(self, student: StudentInput) -> ClassificationResult:
        # Check both score_or_level and explicit exam_score
        num_score = None
        if student.exam_score is not None and str(student.exam_score).strip():
            _, parsed_exam_num, _ = self.parse_score_and_level(student.exam_score)
            if parsed_exam_num is not None:
                num_score = parsed_exam_num

        level, parsed_num, tier = self.parse_score_and_level(student.score_or_level)
        if num_score is None and parsed_num is not None:
            num_score = parsed_num

        # If a numerical exam score was provided and the level was either derived or not explicitly T/H/C
        if num_score is not None:
            # If student.score_or_level was empty or just a number, refine level from score
            str_level = str(student.score_or_level).strip().upper()
            if str_level not in ["T", "H", "C"]:
                level, _, tier = self.parse_score_and_level(num_score)

        level_labels = {
            "T": "Hoàn thành tốt (Mức T)",
            "H": "Hoàn thành (Mức H)",
            "C": "Chưa hoàn thành (Mức C)"
        }
        level_label = level_labels.get(level, "Hoàn thành (Mức H)")

        # Retrieve subject competencies
        subject_key = student.subject
        matched_key = None
        for key in SUBJECT_COMPETENCIES:
            if key.lower() in subject_key.lower() or subject_key.lower() in key.lower():
                matched_key = key
                break

        strengths = []
        growth_areas = []
        if matched_key and matched_key in SUBJECT_COMPETENCIES:
            strengths = SUBJECT_COMPETENCIES[matched_key]["strengths"]
            growth_areas = SUBJECT_COMPETENCIES[matched_key]["growth_areas"]

        # Pedagogical guidance synthesis
        guidance = []
        if num_score is not None:
            guidance.append(f"Điểm bài kiểm tra định kỳ đạt {num_score} điểm.")

        if level == "T":
            guidance.append("Tuyên dương tinh thần học tập xuất sắc, tư duy vững vàng và sự chủ động.")
            if student.teacher_note:
                guidance.append(f"Kết hợp ghi chú của giáo viên: '{student.teacher_note}'.")
            guidance.append("Động viên học sinh tiếp tục duy trì và phát huy năng lực.")
        elif level == "H":
            guidance.append("Ghi nhận sự nỗ lực, hoàn thành yêu cầu bài học.")
            if tier == "H_HIGH":
                guidance.append("Chỉ ra điểm làm tốt và khích lệ nâng cao một kỹ năng cụ thể.")
            else:
                guidance.append("Động viên nhẹ nhàng, chỉ rõ 1 kỹ năng cụ thể cần rèn luyện thêm.")
            if student.teacher_note:
                guidance.append(f"Lồng ghép tinh tế ghi chú: '{student.teacher_note}'.")
        else:  # C
            guidance.append("Dùng ngôn từ ấm áp, khích lệ và hoàn toàn không dùng từ mang tính chê bai.")
            guidance.append("Chỉ rõ biện pháp hỗ trợ/kỹ năng cần rèn thêm hàng ngày.")
            if student.teacher_note:
                guidance.append(f"Chú ý giải quyết khó khăn: '{student.teacher_note}'.")

        return ClassificationResult(
            standardized_level=level,
            level_label=level_label,
            numerical_score=num_score,
            tier=tier,
            subject_focus_strengths=strengths,
            subject_focus_growth=growth_areas,
            guidance_notes=" ".join(guidance)
        )
