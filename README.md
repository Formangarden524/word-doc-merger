# Word Doc Merger

合并指定目录下的 doc 文件，形成一个文件

## 功能

- 合并多个 Word 文档 (.docx) 为一个文件
- 支持 MHT/altChunk 格式（微信公众号导出文档）
- 每个文档以文件名作为标题分隔
- 保留原文档格式

## 依赖安装

```bash
pip3 install python-docx
```

## 使用方法

```bash
python3 merge_docs_with_content.py <输入文件夹> <输出文件>
```

### 示例

```bash
# 合并当前目录下所有 docx 文件
python3 merge_docs_with_content.py ./my_docs ./merged.docx
```

## 特点

- ✅ 支持标准 Word 文档
- ✅ 支持微信公众号导出文档（MHT 格式）
- ✅ 自动提取标题和正文
- ✅ 按文件名排序合并
- ✅ 每个文档之间自动分页

## 作者

Formangarden524
