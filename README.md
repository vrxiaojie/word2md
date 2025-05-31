# 如何使用
## 1. clone本仓库
```shell
git clone https://github.com/vrxiaojie/word2md.git
```
## 2. 安装python依赖
**python版本>=3.8**
```shell
pip install -r requirements.txt
```
## 3.在终端运行程序
```shell
python word2md.py -i InputFile.docx -o OutputFile.md [-l language]
```
**参数解释**

|参数|作用|备注|
|--|--|--|
|-i|输入doc文档名||
|-o|输出markdown文档名||
|-l|统一文档内的代码块语言|可选|

在执行命令后，会提示输入要转换的段落范围。

**注意：** 段落是以word文档中**大纲级别 1级**作为划分依据的
## 4. 检查输出
在输出markdown文档的目录下存放有`images`目录，里面存有该段落所有图片。

