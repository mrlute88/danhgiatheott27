import re
import random
from typing import List, Optional, Set

from utils.tt27_knowledge import synthesize_unique_comment


class OutputValidatorAgent:
    """Agent 3: Validates, sanitizes, deduplicates, and formats comments for Excel with zero-repetition guarantee."""

    def __init__(self, history_window_size: int = 12):
        self.history_window_size = history_window_size
        self.recent_comments: List[str] = []
        # Stores all comments generated for the current class batch (normalized key)
        self.class_comments: Set[str] = set()
        self.class_raw_comments: List[str] = []

    def validate_and_format(
        self,
        raw_comment: str,
        student_name: str,
        standardized_level: str = "H",
        subject: str = "Toán",
        teacher_note: str = "",
        **kwargs
    ) -> str:
        """
        Cleans and normalizes the comment.
        Guarantees non-empty, safe string adhering to length and 100% uniqueness in class.
        """
        if not raw_comment:
            raw_comment = f"Em {student_name} hoàn thành tốt các nội dung học tập và có ý thức tự giác cao."

        # Step 1: Sanitize prefixes, quotes, markdown
        text = self._sanitize(raw_comment)

        # Step 2: Excel formula injection prevention
        text = self._prevent_excel_injection(text)

        # Step 3: Validate length (15 - 35 words target)
        text = self._adjust_length(text, standardized_level)

        # Step 4: Strict Deduplication against the ENTIRE class session
        text = self._ensure_class_wide_diversity(
            text=text,
            student_name=student_name,
            level=standardized_level,
            subject=subject,
            teacher_note=teacher_note
        )

        # Record in class history and sliding window
        norm_key = self._normalize(text)
        self.class_comments.add(norm_key)
        self.class_raw_comments.append(text)

        self.recent_comments.append(text)
        if len(self.recent_comments) > self.history_window_size:
            self.recent_comments.pop(0)

        return text

    def _normalize(self, text: str) -> str:
        """Strips punctuation and lowercases for accurate duplicate detection."""
        return re.sub(r"[^\w\s]", "", text).strip().lower()

    def _is_duplicate_or_too_similar(self, text: str, threshold: float = 0.70) -> bool:
        """Checks if text is identical or shares > 70% word vocabulary with an existing comment in class."""
        norm = self._normalize(text)
        if norm in self.class_comments:
            return True

        words_new = set(norm.split())
        if len(words_new) < 4:
            return False

        for existing_norm in self.class_comments:
            words_exist = set(existing_norm.split())
            union = len(words_new | words_exist)
            intersection = len(words_new & words_exist)
            if union > 0 and (intersection / union) >= threshold:
                return True

        return False

    def _sanitize(self, text: str) -> str:
        # Strip common AI preambles
        patterns = [
            r"^(nhận xét|lời phê|đánh giá)(\s*học sinh)?\s*[:\-\.]\s*",
            r"^học sinh\s*:\s*",
            r"^thưa cô\s*,\s*",
            r"^dưới đây là nhận xét\s*:\s*",
            r"^lời nhận xét\s*:\s*"
        ]
        for pat in patterns:
            text = re.sub(pat, "", text, flags=re.IGNORECASE)

        # Remove surrounding or extraneous quotes
        text = text.strip(" '\"“”«»`*")

        # Remove markdown bold/italic tags
        text = text.replace("**", "").replace("*", "").replace("__", "")

        # Collapse whitespace and newlines to single space
        text = re.sub(r"\s+", " ", text).strip()

        # Capitalize first letter
        if text:
            text = text[0].upper() + text[1:]

        # Ensure ends with punctuation
        if text and text[-1] not in [".", "!", "?"]:
            text += "."

        return text

    def _prevent_excel_injection(self, text: str) -> str:
        """Prevents leading '=', '+', '-', '@' from executing as an Excel formula."""
        if text.startswith(("=", "+", "-", "@")):
            text = "'" + text
        return text

    def _adjust_length(self, text: str, level: str) -> str:
        words = text.split()
        word_count = len(words)

        # If too short (< 12 words), add supportive closing phrase
        if word_count < 12:
            closings = {
                "T": "Em hãy tiếp tục phát huy tinh thần học tập tích cực này nhé.",
                "H": "Cần cố gắng rèn luyện thêm để đạt kết quả cao hơn.",
                "C": "Em hãy cố gắng nhiều hơn mỗi ngày, thầy cô luôn hỗ trợ em."
            }
            extra = closings.get(level, "Em tiếp tục cố gắng nhé.")
            text = f"{text} {extra}"

        # If too long (> 38 words), split sentences and keep first 2
        elif word_count > 38:
            sentences = re.split(r"(?<=[.!?])\s+", text)
            if len(sentences) >= 2:
                text = f"{sentences[0]} {sentences[1]}"
            else:
                text = " ".join(words[:32]) + "."

        return text

    def _ensure_class_wide_diversity(
        self,
        text: str,
        student_name: str,
        level: str,
        subject: str,
        teacher_note: str
    ) -> str:
        """Guarantees the text is completely non-duplicating across the entire classroom."""
        if not self._is_duplicate_or_too_similar(text):
            return text

        # Attempt 1: Synonym replacement
        replacements = [
            ("Em nắm vững kiến thức", "Có kỹ năng tiếp thu bài tốt"),
            ("Em hoàn thành tốt", "Nắm chắc các nội dung học tập"),
            ("Tư duy logic tốt", "Có khả năng lập luận và tư duy nhanh"),
            ("Cần rèn luyện thêm", "Nên chú ý luyện tập nhiều hơn"),
            ("Cần cố gắng", "Hãy tích cực rèn luyện thêm"),
            ("Em đọc to, rõ ràng", "Kỹ năng đọc lưu loát, to rõ"),
            ("Chữ viết nắn nót", "Trình bày bài sạch đẹp, cẩn thận"),
            ("Tính toán nhanh", "Kỹ năng làm tính nhanh nhẹn"),
            ("rất chính xác", "chuẩn xác và cẩn thận"),
            ("tiếp thu nhanh", "nắm bắt bài học rất nhanh"),
            ("tích cực phát biểu", "hăng hái xây dựng bài"),
            ("Cần tiếp tục phát huy.", "Thầy/Cô rất khen ngợi em."),
            ("Thầy/Cô rất khen ngợi em.", "Hãy duy trì phong độ này nhé.")
        ]
        
        mutated = text
        for old, new in replacements:
            if old in mutated:
                mutated = mutated.replace(old, new, 1)
                if not self._is_duplicate_or_too_similar(mutated):
                    return mutated

        # Attempt 2: If still duplicate or too similar, synthesize an unused, unique comment
        fresh_comment = synthesize_unique_comment(
            student_name=student_name,
            subject=subject,
            level=level,
            teacher_note=teacher_note,
            used_comments=self.class_comments,
            allow_name=True
        )

        return self._sanitize(fresh_comment)

    def reset_history(self):
        """Resets the history for a new class or file processing batch."""
        self.recent_comments.clear()
        self.class_comments.clear()
        self.class_raw_comments.clear()

