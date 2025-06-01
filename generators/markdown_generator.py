from converters import MarkdownConverter
from processors import ImageProcessor
from typing import List, Dict, Any, Tuple
from docx import Document


class MarkdownGenerator:
    """Markdown生成器"""

    def __init__(self, converter: MarkdownConverter, image_processor: ImageProcessor):
        """
        初始化Markdown生成器

        Args:
            converter: Markdown转换器
            image_processor: 图片处理器
        """
        self.converter = converter
        self.image_processor = image_processor

    def _process_paragraph_runs(self, para, rels, section_index: int,
                                image_count: int)-> Tuple[str, List[str], int]:
        """
        处理段落中的runs

        Args:
            para: 段落对象
            rels: 文档关系
            section_index: 章节索引
            image_count: 图片计数

        Returns:
            (合并的文本, 图片markdown行列表, 更新后的图片计数)
        """
        merged_text = ""
        current_bold = None
        image_lines = []

        for run in para.runs:
            # 处理图片
            img_lines, image_count = self.image_processor.process_images_in_run(
                run, rels, section_index, image_count
            )
            image_lines.extend(img_lines)

            # 处理文本和格式
            text = run.text
            bold = run.bold

            if bold != current_bold:
                if current_bold is True:
                    merged_text += "**"
                if bold is True:
                    merged_text += "**"
                current_bold = bold

            merged_text += text

        if current_bold is True:
            merged_text += "**"

        return merged_text, image_lines, image_count

    def generate_markdown_for_sections(self, sections: List[Dict[str, Any]],
                                       document: Document) -> List[str]:
        """
        为选定章节生成Markdown

        Args:
            sections: 选定的章节列表
            document: Word文档对象

        Returns:
            Markdown行列表
        """
        markdown_lines = []
        rels = document.part._rels

        for sec in sections:
            image_count = 1
            sec_index = sec["index"]
            markdown_lines.append(f"## {sec['title']}")

            in_code_block = False

            for para in sec["content"]:
                style = para.style.name.strip()
                para_text = para.text.strip()

                # 处理标题
                if style == "Heading 2":
                    if in_code_block:
                        markdown_lines.append("```")
                        in_code_block = False
                    markdown_lines.append(f"### {para_text}")
                    continue
                elif style == "Heading 3":
                    if in_code_block:
                        markdown_lines.append("```")
                        in_code_block = False
                    markdown_lines.append(f"#### {para_text}")
                    continue

                # 处理代码块
                if style == "Code":
                    if not in_code_block:
                        markdown_lines.append("```" + self.converter.code_language)
                        in_code_block = True
                    markdown_lines.append(para.text)
                    continue
                else:
                    if in_code_block:
                        markdown_lines.append("```")
                        in_code_block = False

                # 处理正文、图片和格式
                merged_text, image_lines, image_count = self._process_paragraph_runs(
                    para, rels, sec_index, image_count
                )

                # 添加图片行
                markdown_lines.extend(image_lines)

                # 添加文本行
                if merged_text.strip():
                    converted_text = self.converter.convert_plain_urls_to_md(merged_text.strip())
                    markdown_lines.append(converted_text)

                markdown_lines.append("")

            if in_code_block:
                markdown_lines.append("```")

        return markdown_lines