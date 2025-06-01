import os

from converters import MarkdownConverter
from generators import MarkdownGenerator
from processors import ImageProcessor
from selector import SectionSelector
from parsers import WordDocumentParser
from typing import List, Optional, Set, Union


class Word2MarkdownConverter:
    """Word到Markdown转换器主类"""

    def __init__(self, code_language: str = ""):
        """
        初始化转换器

        Args:
            code_language: 代码块语言标识
        """
        self.converter = MarkdownConverter(code_language)
        self.selector = SectionSelector()

    def convert(self, docx_path: str, output_md_path: str,
                section_range: Optional[Union[Set[int], List[int]]] = None) -> bool:
        """
        执行转换

        Args:
            docx_path: 输入Word文档路径
            output_md_path: 输出Markdown文件路径
            section_range: 要转换的章节范围（可选）

        Returns:
            转换是否成功
        """
        try:
            # 解析文档
            parser = WordDocumentParser(docx_path)

            # 选择章节
            if section_range is None:
                selected_sections = self.selector.select_sections_interactive(parser)
            else:
                selected_sections = self.selector.select_sections_by_range(parser, section_range)

            if not selected_sections:
                print("没有选择任何章节")
                return False

            # 准备输出目录和图片处理器
            output_dir = os.path.dirname(os.path.abspath(output_md_path))
            image_processor = ImageProcessor(output_dir)

            # 生成Markdown
            generator = MarkdownGenerator(self.converter, image_processor)
            markdown_lines = generator.generate_markdown_for_sections(
                selected_sections, parser.document
            )

            # 写入文件
            with open(output_md_path, "w", encoding="utf-8") as f:
                f.write("\n".join(markdown_lines))

            print(f"\n✅ 转换完成：{output_md_path}")
            print(f"🖼️ 图片输出目录：{image_processor.image_dir}")

            return True

        except Exception as e:
            print(f"转换过程中发生错误：{e}")
            return False