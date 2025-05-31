import argparse
from docx import Document
import os
import shutil
from docx.oxml.ns import qn

def word_to_markdown(docx_path, output_md_path, code_lang, section_range=None):
    # 将图片与输出文件存放在同一个目录下
    output_dir = os.path.dirname(os.path.abspath(output_md_path))
    image_dir = os.path.join(output_dir, "images")
    doc = Document(docx_path)

    # 按 Heading 1 分段
    sections = []
    current_section = {"title": "", "content": [], "index": 0}
    section_index = 0

    for para in doc.paragraphs:
        style = para.style.name.strip()
        if style == "Heading 1":
            if current_section["content"]:
                sections.append(current_section)
            section_index += 1
            current_section = {
                "title": para.text.strip(),
                "content": [],
                "index": section_index
            }
        else:
            current_section["content"].append(para)
    if current_section["content"]:
        sections.append(current_section)

    # 用户选择段落
    if section_range is None:
        print("检测到以下段落：")
        for sec in sections:
            print(f"{sec['index']}: {sec['title']}")
        user_input = input("请输入要转换的段落范围（例如 2-4 或 1,3,5）：").strip()
        if "-" in user_input:
            start, end = map(int, user_input.split("-"))
            selected_sections = [s for s in sections if start <= s["index"] <= end]
        else:
            indices = set(map(int, user_input.split(",")))
            selected_sections = [s for s in sections if s["index"] in indices]
    else:
        selected_sections = [s for s in sections if s["index"] in section_range]

    # 清空图片目录
    if os.path.exists(image_dir):
        shutil.rmtree(image_dir)
    os.makedirs(image_dir)

    # 图片使用 doc.part.rels 来定位 ID 和图片数据
    rels = doc.part._rels

    markdown_lines = []

    for sec in selected_sections:
        image_count = 1
        sec_index = sec["index"]
        markdown_lines.append(f"## {sec['title']}")

        in_code_block = False
        in_info_block = False

        for para in sec["content"]:
            style = para.style.name.strip()
            para_text = para.text.strip()

            if style == "Heading 2":
                markdown_lines.append(f"### {para_text}")
                continue
            elif style == "Heading 3":
                markdown_lines.append(f"#### {para_text}")
                continue

            if style == "Code":
                if not in_code_block:
                    markdown_lines.append("```" + code_lang)
                    in_code_block = True
                markdown_lines.append(para.text)
                continue
            else:
                if in_code_block:
                    markdown_lines.append("```")
                    in_code_block = False

            # 正文 + 图片 + 加粗处理
            merged_text = ""
            current_bold = None

            for run in para.runs:
                # 判断是否为图片
                for drawing in run._element.iter():
                    if drawing.tag.endswith('drawing'):
                        blip_elems = drawing.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                        for blip in blip_elems:
                            embed = blip.get(qn("r:embed"))
                            if embed in rels:
                                image_part = rels[embed]._target
                                image_data = image_part.blob
                                image_name = f"{sec_index}-{image_count}.jpg"
                                image_path = os.path.join(image_dir, image_name)
                                with open(image_path, "wb") as f:
                                    f.write(image_data)
                                markdown_lines.append(f"![img{sec_index}-{image_count}](./{image_dir}/{image_name})")
                                image_count += 1
                        break  # 如果已找到一个图片，不再继续搜索该 run

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

            if merged_text.strip():
                markdown_lines.append(merged_text.strip())

            markdown_lines.append("")

        if in_code_block:
            markdown_lines.append("```")

    with open(output_md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(markdown_lines))

    print(f"\n✅ 转换完成：{output_md_path}")
    print(f"🖼️ 图片输出目录：{image_dir}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="将 Word 文件按段落转换为 Markdown")
    parser.add_argument("-i", "--input", required=True, help="输入的 .docx 文件路径")
    parser.add_argument("-o", "--output", required=True, help="输出的 .md 文件路径")
    parser.add_argument("-l", "--lang", required=False, default="", help="代码块语言类型")
    args = parser.parse_args()

    word_to_markdown(args.input, args.output, args.lang)
