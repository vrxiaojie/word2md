import argparse
from converters import Word2MarkdownConverter

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