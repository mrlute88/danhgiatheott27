"""
commentator.py - Agent 2: TT27 Pedagogical Commentator (Gemini Powered)
Generates personalized, constructive, encouraging primary school comments complying with Circular 27.
Includes retry logic, rate limit handling, and robust offline rule-based fallback.
"""

import time
import random
import re
from typing import Optional, Tuple, Set, List
import google.generativeai as genai
from google.api_core.exceptions import ResourceExhausted, GoogleAPIError

from .classifier import StudentInput, ClassificationResult
from utils.tt27_knowledge import (
    TT27_COMMENT_BANK,
    SUBJECT_COMPETENCIES,
    synthesize_unique_comment,
)


SYSTEM_PROMPT_TEMPLATE = """Bạn là một chuyên gia giáo dục tiểu học tại Việt Nam, am hiểu sâu sắc Thông tư 27/2020/TT-BGDĐT.
Nhiệm vụ của bạn là viết nhận xét học sinh ngắn gọn, tinh tế, mang tính khuyến khích và cá nhân hóa cao.

Thông tin học sinh:
- Họ và tên: {student_name}
- Môn học: {subject} - Lớp: {grade}
- Mức độ hoàn thành / Điểm số: {score_or_level} (Mức chuẩn hóa: {level_label})
- Ghi chú từ giáo viên: {teacher_note}
- Thời điểm đánh giá: {period}
- Phong cách: {tone}
- Chỉ dẫn sư phạm bổ sung: {guidance_notes}

YÊU CẦU QUAN TRỌNG VỀ ĐỘC BẢN VÀ CHỐNG TRÙNG LẶP CHO LỚP HỌC:
Các học sinh trong cùng một lớp TUYỆT ĐỐI KHÔNG ĐƯỢC có nhận xét giống hệt nhau hoặc lặp lại cấu trúc rập khuôn.
Dưới đây là một số lời nhận xét đã được viết cho các học sinh trước đó trong lớp:
{used_sample_str}
-> Bạn BẮT BUỘC phải dùng cách mở đầu khác, chọn khía cạnh năng lực hoặc phẩm chất khác của môn học, và lời động viên KHÁC BIỆT so với các câu trên.

Quy tắc viết nhận xét:
1. Độ dài: 1 đến 2 câu (khoảng 15-32 từ).
2. Tốt/Điểm cao (9-10): Tuyên dương ưu điểm nổi bật (tư duy tốt, cẩn thận, sáng tạo) và khuyến khích phát huy.
3. Hoàn thành/Điểm trung bình (7-8): Ghi nhận sự nỗ lực, nêu rõ kỹ năng làm tốt và chỉ ra 1 điểm cần luyện thêm nhẹ nhàng.
4. Chưa hoàn thành/Điểm yếu (<7): Dùng từ ngữ nhẹ nhàng, mang tính hỗ trợ, không dùng từ tiêu cực. Chỉ rõ kỹ năng cụ thể cần phụ đạo/rèn thêm.
5. Không viết chung chung kiểu "Học tốt, ngoan". Hãy gắn với kỹ năng môn học (VD: "Tính toán nhanh", "Đọc diễn cảm", "Giữ vở sạch chữ đẹp").
6. Chỉ trả về duy nhất nội dung nhận xét, không kèm lời mở đầu hay giải thích.
"""


class TT27CommentatorAgent:
    """Agent 2: Generates Circular 27 compliant student comments using Gemini API or fallback."""

    def __init__(self, api_key: Optional[str] = None, model_name: str = "gemini-3.6-flash"):
        self.api_key = api_key
        self.model_name = model_name
        self.is_configured = False
        if api_key and api_key.strip():
            self._configure(api_key.strip())

    def _configure(self, key: str):
        try:
            genai.configure(api_key=key)
            self.is_configured = True
        except Exception:
            self.is_configured = False

    @staticmethod
    def get_available_models(api_key: str) -> List[str]:
        """Fetches active generative models supported: locked to gemini-3.6-flash."""
        return ["gemini-3.6-flash"]

    @staticmethod
    def test_api_key(api_key: str, model_name: str = "gemini-3.6-flash") -> Tuple[bool, str]:
        """Tests validity of the provided Gemini API Key and model name."""
        if not api_key or not api_key.strip():
            return False, "Vui lòng nhập Google Gemini API Key."
        try:
            genai.configure(api_key=api_key.strip())
            model = genai.GenerativeModel(model_name)
            response = model.generate_content("Chào bạn, hãy phản hồi 'OK' nếu bạn sẵn sàng.")
            if response and response.text:
                return True, f"Kết nối Gemini API thành công! Mô hình: {model_name}"
            return False, "Không nhận được phản hồi từ Gemini API."
        except Exception as e:
            err_msg = str(e)
            if "API_KEY_INVALID" in err_msg or "400" in err_msg:
                return False, "API Key không hợp lệ. Vui lòng kiểm tra lại khóa API của bạn."
            elif "RESOURCE_EXHAUSTED" in err_msg or "429" in err_msg:
                return False, "API Key đã vượt hạn mức yêu cầu (Rate Limit/Quota). Thử lại sau ít phút."
            elif "404" in err_msg or "not found" in err_msg.lower() or "no longer available" in err_msg.lower():
                return False, f"Mô hình '{model_name}' không tìm thấy trên API key này (Lỗi 404). Hãy đảm bảo dùng model 'gemini-3.6-flash'."
            return False, f"Lỗi kết nối: {err_msg}"

    def generate_comment(
        self,
        student: StudentInput,
        classification: ClassificationResult,
        recent_comments: Optional[list] = None,
        used_comments: Optional[set] = None,
        **kwargs
    ) -> str:
        """
        Generates a 100% non-duplicating comment for a student in the classroom.
        Falls back to combinatorial synthesizer if API fails, is not configured, or generates a duplicate.
        """
        if used_comments is None:
            used_comments = kwargs.get("used_comments") or set()

        if self.is_configured and self.api_key:
            try:
                comment = self._call_gemini_with_retry(student, classification, recent_comments)
                if comment and len(comment.strip()) > 10:
                    clean_norm = re.sub(r"[^\w\s]", "", comment).lower()
                    if clean_norm not in used_comments:
                        used_comments.add(clean_norm)
                        return comment.strip()
            except Exception:
                pass

        # Offline / Fallback generation guaranteed 100% zero duplicate
        return self._generate_fallback(student, classification, used_comments)

    def _call_gemini_with_retry(
        self,
        student: StudentInput,
        classification: ClassificationResult,
        recent_comments: Optional[list] = None,
        max_retries: int = 3
    ) -> str:
        """Calls Gemini API with exponential backoff on 429/quota errors."""
        used_sample_str = "Chưa có nhận xét nào trước đó."
        if recent_comments:
            last_few = recent_comments[-6:]
            used_sample_str = "\n".join([f"- \"{c}\"" for c in last_few])

        prompt = SYSTEM_PROMPT_TEMPLATE.format(
            student_name=student.student_name,
            subject=student.subject,
            grade=student.grade,
            score_or_level=student.score_or_level,
            level_label=classification.level_label,
            teacher_note=student.teacher_note if student.teacher_note else "Không có ghi chú riêng",
            period=student.period,
            tone=student.tone,
            guidance_notes=classification.guidance_notes,
            used_sample_str=used_sample_str
        )

        model = genai.GenerativeModel(
            model_name=self.model_name,
            generation_config={
                "temperature": 0.90,
                "top_p": 0.95,
                "max_output_tokens": 150,
            }
        )

        delay = 2.0
        for attempt in range(max_retries):
            try:
                response = model.generate_content(prompt)
                if response and response.text:
                    return response.text.strip()
            except ResourceExhausted:
                if attempt < max_retries - 1:
                    time.sleep(delay)
                    delay *= 2
                else:
                    raise
            except Exception as e:
                err_str = str(e)
                # Auto-recovery using gemini-3.6-flash
                if ("404" in err_str or "no longer available" in err_str.lower() or "not found" in err_str.lower()):
                    try:
                        self.model_name = "gemini-3.6-flash"
                        model = genai.GenerativeModel("gemini-3.6-flash")
                        response = model.generate_content(prompt)
                        if response and response.text:
                            return response.text.strip()
                    except Exception:
                        pass
                if attempt < max_retries - 1:
                    time.sleep(1.0)
                else:
                    raise

        return ""

    def _generate_fallback(
        self,
        student: StudentInput,
        classification: ClassificationResult,
        used_comments: Optional[set] = None
    ) -> str:
        """Rule-based combinatorial fallback generator guaranteeing zero duplicate comments across class."""
        return synthesize_unique_comment(
            student_name=student.student_name,
            subject=student.subject,
            level=classification.standardized_level,
            tier=classification.tier,
            teacher_note=student.teacher_note or "",
            tone=student.tone,
            used_comments=used_comments,
            allow_name=True
        )
