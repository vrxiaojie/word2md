import argparse
import re
import os
import shutil
from typing import List, Dict, Any, Optional, Set, Union, Tuple
from docx import Document
from docx.oxml.ns import qn


class MarkdownConverter:
    """Word文档到Markdown转换器"""

    def __init__(self, code_language: str = ""):
        """
        初始化转换器

        Args:
            code_language: 代码块的默认语言标识
        """
        self.code_language = code_language
        self.image_counter = 1

    @staticmethod
    def convert_plain_urls_to_md(text: str) -> str:
        """
        将纯文本URL转换为Markdown链接格式

        Args:
            text: 包含URL的文本

        Returns:
            转换后的Markdown文本
        """
        url_pattern = r"(https?://[^\s\)\]\}<>\"\'，。：；！、]+)"

        def replacer(match):
            url = match.group(1)
            return f"[{url}]({url})"

        return re.sub(url_pattern, replacer, text)


class WordDocumentParser:
    """Word文档解析器"""

    def __init__(self, docx_path: str):
        """
        初始化文档解析器

        Args:
            docx_path: Word文档路径
        """
        self.docx_path = docx_path
        self.document = Document(docx_path)
        self.sections = []
        self._parse_sections()

    def _parse_sections(self):
        """解析文档中的章节"""
        current_section = {"title": "", "content": [], "index": 0}
        section_index = 0

        for para in self.document.paragraphs:
            style = para.style.name.strip()
            if style == "Heading 1":
                if current_section["content"]:
                    self.sections.append(current_section)
                section_index += 1
                current_section = {
                    "title": para.text.strip(),
                    "content": [],
                    "index": section_index
                }
            else:
                current_section["content"].append(para)

        if current_section["content"]:
            self.sections.append(current_section)

    def get_sections(self) -> List[Dict[str, Any]]:
        """获取所有章节"""
        return self.sections

    def get_section_titles(self) -> Dict[int, str]:
        """获取章节标题映射"""
        return {sec["index"]: sec["title"] for sec in self.sections if sec["index"] > 0}

    def get_max_section_index(self) -> int:
        """获取最大章节索引"""
        return max((sec["index"] for sec in self.sections), default=0)


class SectionSelector:
    """章节选择器"""

    @staticmethod
    def parse_range_input(user_input: str) -> Set[int]:
        """
        解析用户输入的范围

        Args:
            user_input: 用户输入（如 "2-4" 或 "1,3,5"）

        Returns:
            选中的章节索引集合
        """
        user_input = user_input.strip()
        if "-" in user_input:
            start, end = map(int, user_input.split("-"))
            return set(range(start, end + 1))
        else:
            return set(map(int, user_input.split(",")))

    @staticmethod
    def validate_selection(indices: Set[int], max_index: int) -> bool:
        """
        验证选择是否有效

        Args:
            indices: 选中的索引
            max_index: 最大有效索引

        Returns:
            是否有效
        """
        return all(1 <= idx <= max_index for idx in indices)

    def select_sections_interactive(self, parser: WordDocumentParser) -> List[Dict[str, Any]]:
        """
        交互式选择章节

        Args:
            parser: 文档解析器

        Returns:
            选中的章节列表
        """
        sections = parser.get_sections()
        section_titles = parser.get_section_titles()
        max_index = parser.get_max_section_index()

        print("检测到以下段落：")
        for idx, title in section_titles.items():
            print(f"{idx}: {title}")

        while True:
            user_input = input("请输入要转换的段落范围（例如 2-4 或 1,3,5）：").strip()
            try:
                selected_indices = self.parse_range_input(user_input)
                if self.validate_selection(selected_indices, max_index):
                    return [s for s in sections if s["index"] in selected_indices]
                else:
                    print(f"错误，请检查输入的段落数值！最大值为{max_index}")
            except ValueError:
                print("输入格式错误，请重新输入！")

    def select_sections_by_range(self, parser: WordDocumentParser,
                                 section_range: Union[Set[int], List[int]]) -> List[Dict[str, Any]]:
        """
        根据指定范围选择章节

        Args:
            parser: 文档解析器
            section_range: 章节范围

        Returns:
            选中的章节列表
        """
        if isinstance(section_range, list):
            section_range = set(section_range)

        sections = parser.get_sections()
        return [s for s in sections if s["index"] in section_range]


class ImageProcessor:
    """图片处理器"""

    def __init__(self, output_dir: str):
        """
        初始化图片处理器

        Args:
            output_dir: 输出目录
        """
        self.output_dir = output_dir
        self.image_dir = os.path.join(output_dir, "images")
        self._prepare_image_directory()

    def _prepare_image_directory(self):
        """准备图片目录"""
        if os.path.exists(self.image_dir):
            shutil.rmtree(self.image_dir)
        os.makedirs(self.image_dir)

    def process_images_in_run(self, run, rels, section_index: int,
                              image_count: int) -> Tuple[List[str], int]:
        """
        处理run中的图片

        Args:
            run: Word文档中的run对象
            rels: 文档关系
            section_index: 章节索引
            image_count: 当前图片计数

        Returns:
            (markdown图片行列表, 更新后的图片计数)
        """
        markdown_lines = []

        for drawing in run._element.iter():
            if drawing.tag.endswith('drawing'):
                blip_elems = drawing.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                for blip in blip_elems:
                    embed = blip.get(qn("r:embed"))
                    if embed in rels:
                        image_part = rels[embed]._target
                        image_data = image_part.blob
                        image_name = f"{section_index}-{image_count}.jpg"
                        image_path = os.path.join(self.image_dir, image_name)

                        with open(image_path, "wb") as f:
                            f.write(image_data)

                        markdown_lines.append(f"![img{section_index}-{image_count}]({self.image_dir}/{image_name})")
                        image_count += 1
                break

        return markdown_lines, image_count


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


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="将 Word 文件按段落转换为 Markdown")
    parser.add_argument("-i", "--input", required=True, help="输入的 .docx 文件路径")
    parser.add_argument("-o", "--output", required=True, help="输出的 .md 文件路径")
    parser.add_argument("-l", "--lang", required=False, default="", help="代码块语言类型")
    args = parser.parse_args()

    converter = Word2MarkdownConverter(args.lang)
    converter.convert(args.input, args.output)


if __name__ == "__main__":
    main()