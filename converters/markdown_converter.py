import re


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

