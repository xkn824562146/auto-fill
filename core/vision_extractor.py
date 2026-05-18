"""AI 视觉提取器：将 PDF 页面转为图片，调用 AI 模型提取指定字段。"""

import logging
import base64
import json
import re
import fitz
from openai import OpenAI

logger = logging.getLogger(__name__)


class VisionExtractor:
    """通过 AI 视觉模型从 PDF 图片中提取指定字段的值。"""

    def __init__(self, config: dict, aliases: dict[str, list[str]] | None = None):
        self.client = OpenAI(
            base_url=config["base_url"],
            api_key=config["api_key"],
        )
        self.model = config["model"]
        # 别名配置的原始映射：alias_key -> [alias_values]
        self.aliases: dict[str, list[str]] = aliases or {}
        # 构建反向映射：别名 -> alias_key
        self.alias_to_field: dict[str, str] = {}
        if aliases:
            for field_name, alias_list in aliases.items():
                for alias in alias_list:
                    self.alias_to_field[alias] = field_name

    def extract(self, pdf_path: str, field_names: list[str], dpi: int = 300) -> dict[str, str | None]:
        """从 PDF 中提取指定字段的值。

        Args:
            pdf_path: PDF 文件路径
            field_names: 需要提取的字段名列表
            dpi: 渲染分辨率

        Returns:
            {字段名: 提取值} 字典，找不到的字段值为 None
        """
        images = self._pdf_to_images(pdf_path, dpi)
        all_results: dict[str, str | None] = {name: None for name in field_names}

        for page_idx, img_b64 in enumerate(images):
            page_result = self._extract_from_page(img_b64, field_names, page_idx + 1)
            # 合并结果：只更新尚未找到的字段
            for key, value in page_result.items():
                if value is not None and all_results.get(key) is None:
                    all_results[key] = value

            # 如果所有字段都找到了，提前退出
            if all(v is not None for v in all_results.values()):
                break

        return all_results

    def _pdf_to_images(self, pdf_path: str, dpi: int) -> list[str]:
        """将 PDF 每页转为 base64 编码的 PNG 图片。"""
        doc = fitz.open(pdf_path)
        images = []
        scale = dpi / 72

        for page in doc:
            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat)
            img_bytes = pix.tobytes("png")
            img_b64 = base64.b64encode(img_bytes).decode("utf-8")
            images.append(img_b64)

        doc.close()
        return images

    def _extract_from_page(self, img_b64: str, field_names: list[str], page_num: int) -> dict[str, str | None]:
        """从单页图片中提取字段的原始测量值。"""
        fields_str = "、".join(field_names)

        # 构建别名提示（支持字段名包含 alias_key 的情况，如"内焰尖高度FS"包含"内焰尖高度"）
        alias_hint = ""
        if self.aliases:
            lines = []
            for alias_key, alias_list in self.aliases.items():
                # 检查是否有字段名包含此 alias_key
                matched_field = next((f for f in field_names if alias_key in f), None)
                if matched_field:
                    lines.append(f'"{matched_field}" 也可能写作 "{"、".join(alias_list)}"')
            if lines:
                alias_hint = "\n7. 特别注意以下别名对应关系：" + "；".join(lines)

        prompt = f"""这是一份检验报告的第 {page_num} 页照片。请寻找以下检测项目的实际测量值：[{fields_str}]

要求：
1. 以 JSON 格式返回，key 为检测项名称，value 为实际测量数值
2. 找不到的字段返回 null
3. 请进行模糊语义匹配，例如"导热系数"可能写作"λ值"、"热导率"或"导热系数（干态）"
4. 只提取测量数值，不要提取标准要求（如≤0.16）或合格判定
5. 数值请保留原始精度，不要四舍五入
6. 只返回 JSON，不要其他文字{alias_hint}

示例返回格式：
{{"导热系数（干态）": "0.035", "炉内温升": "18.2", "持续燃烧时间": "0"}}"""

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/png;base64,{img_b64}",
                                },
                            },
                        ],
                    }
                ],
                temperature=0.1,
            )

            content = response.choices[0].message.content
            logger.info(
                "[VisionExtractor] 第 %s 页 AI 原始返回:\n%s",
                page_num, content[:500],
            )
            return self._parse_response(content, field_names)

        except Exception as e:
            logger.error("[VisionExtractor] 第 %s 页调用 AI 失败: %s", page_num, e)
            return {name: None for name in field_names}

    def _parse_response(self, content: str, field_names: list[str]) -> dict[str, str | None]:
        """解析 AI 返回的 JSON，容错处理。"""
        result = {name: None for name in field_names}

        # 尝试从回复中提取 JSON（兼容 markdown 代码块）
        json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
        if json_match:
            json_str = json_match.group(1)
        else:
            json_match = re.search(r'\{[^{}]*\}', content, re.DOTALL)
            json_str = json_match.group() if json_match else None

        if not json_str:
            return result

        try:
            parsed = json.loads(json_str)
        except json.JSONDecodeError:
            return result

        # 提取"检验结果"值（AI 通常会单独返回这个 key）
        result_value = None
        for key in ("检验结果", "检测结果"):
            if key in parsed and parsed[key] is not None:
                result_value = str(parsed[key])
                break

        # 做模糊匹配，将 AI 返回的 key 映射到 field_names
        from thefuzz import process

        for ai_key, ai_value in parsed.items():
            if ai_value is None:
                continue
            # 跳过非检测项 key（如"以下空白"、纯数字等）
            if ai_key in ("检验结果", "检测结果", "以下空白", "以下为空白"):
                continue
            if re.match(r'^\d+$', ai_key):
                continue
            # 精确匹配
            if ai_key in result:
                result[ai_key] = str(ai_value)
                continue
            # 别名匹配：AI 返回的 key 是某个字段的别名（支持字段名包含 alias_key）
            if ai_key in self.alias_to_field:
                alias_key = self.alias_to_field[ai_key]
                target = next((f for f in result if alias_key in f), None)
                if target and result[target] is None:
                    result[target] = str(ai_value)
                    continue
            # 模糊匹配
            match, score = process.extractOne(ai_key, field_names)
            if score >= 60:
                if result[match] is None:
                    result[match] = str(ai_value)

        # 如果有"检验结果"值：替换含逗号的原始数据，或填充未匹配的字段
        if result_value is not None:
            for key in result:
                if result[key] is None:
                    result[key] = result_value
                elif ',' in result[key]:
                    result[key] = result_value

        return result
