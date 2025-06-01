import os
import shutil
from docx.oxml.ns import qn
from typing import List, Tuple


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