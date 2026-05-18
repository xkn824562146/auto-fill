"""语义映射层：将 AI 提取结果映射到 Word 模板字段。"""

from thefuzz import process


class FieldMapper:
    """将 AI 返回的字段名映射到 Word 模板的标签名。"""

    def map_fields(self, ai_result: dict[str, str | None], word_fields: list[str], threshold: int = 60) -> dict[str, str | None]:
        """将 AI 结果映射到 Word 字段。

        Args:
            ai_result: AI 提取的 {字段名: 值} 字典
            word_fields: Word 模板中的标签名列表
            threshold: 模糊匹配阈值 (0-100)

        Returns:
            {Word 标签名: 值} 字典
        """
        mapped: dict[str, str | None] = {f: None for f in word_fields}

        for ai_key, ai_value in ai_result.items():
            if ai_value is None:
                continue

            # 精确匹配
            if ai_key in mapped and mapped[ai_key] is None:
                mapped[ai_key] = ai_value
                continue

            # 模糊匹配
            match, score = process.extractOne(ai_key, word_fields)
            if score >= threshold and mapped[match] is None:
                mapped[match] = ai_value

        return mapped
