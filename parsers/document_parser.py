from docx import Document
from typing import List, Dict, Any


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
