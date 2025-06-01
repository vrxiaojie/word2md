from typing import List, Dict, Any, Set, Union
from parsers import WordDocumentParser


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

    @staticmethod
    def select_sections_by_range(parser: WordDocumentParser,
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






