# cn-gov-docx

[![PyPI version](https://img.shields.io/pypi/v/cn-gov-docx.svg)](https://pypi.org/project/cn-gov-docx/)
[![Python](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)

一个基于 `python-docx` 的中文公文 Word（.docx）生成工具，用于快速生成符合政府机关常见排版规范的文档结构。

本项目目标是提供一套**稳定、可复用的公文排版生成层**，减少重复排版工作。

---

## 功能特点

- 自动生成 A4 公文版式（页边距已预设）
- 内置中文公文常用字体规则
- 支持标题 / 分级标题 / 正文结构
- 自动处理首行缩进与行距
- 封装常用 Word 样式设置逻辑
- 基于 `python-docx`，无复杂依赖

---

## 安装

```bash
pip install cn-gov-docx
````

---

## 快速开始

```python
from cn_gov_docx import CnGovDocx

doc = CnGovDocx()

doc.add_title("关于推进重点工作的通知")

doc.add_h1("一、总体要求")
doc.add_text("这里是正文内容，用于测试公文排版效果。")

doc.add_h2("（一）工作目标")
doc.add_text("继续补充正文内容。")

doc.save("output.docx")
```

---

## 默认排版规则

### 页面设置（A4）

* 上边距：3.7 cm
* 下边距：3.5 cm
* 左边距：2.8 cm
* 右边距：2.6 cm

---

### 字体规范

| 类型         | 字体        |
| ---------- | --------- |
| 标题         | 方正小标宋简体   |
| 一级/二级/三级标题 | 黑体        |
| 正文         | 仿宋_GB2312 |

---

### 正文格式

* 字号：16pt（接近三号）
* 行距：28pt
* 首行缩进：2字符
* 对齐方式：两端对齐

---

## API

### `GovDocx()`

创建文档实例。

```python
doc = GovDocx()
```

---

### `add_title(text)`

添加居中主标题。

```python
doc.add_title("标题")
```

---

### `add_h1(text)`

添加一级标题，同样的还有二、三级标题。

```python
doc.add_h1("一级标题")
doc.add_h2("二级标题")
doc.add_h3("三级标题")
```

---

### `add_text(text, indent_first_line=True, remove_newlines=True)`

添加正文内容。

```python
doc.add_text("正文内容")
```

#### 参数

* `indent_first_line`：是否首行缩进（默认 `True`）
* `remove_newlines`：是否移除换行并合并为单段（默认 `True`）


#### 针对`remove_newlines`的说明

Python 三引号换行通常用于代码可读性，不代表真实段落结构。该参数用于区分“格式化换行”与“真实排版结构”。

##### `True`（默认）

* 忽略所有换行符（含 `"""..."""` 内换行）
* 合并为单一段落
* 用于公文正文标准排版

##### `False`

* 保留原始换行结构
* 每行按原样写入文档
* 可能形成多段或显式换行

---

### `save(path)`

保存文件。

```python
doc.save("result.docx")
```

---

## 使用说明（重要）

### 1. 字体依赖

本项目依赖系统字体：

* 方正小标宋简体
* 黑体
* 仿宋_GB2312

如果系统缺失字体，Word 会自动替换字体，可能导致排版不一致。

---

### 2. 适用范围

适用于：

* 政府公文格式草稿
* 国企/事业单位文书
* 标准化报告生成
* 批量文档生成场景

不适用于：

* 精确符合某地区红头文件模板（需二次定制）
* 复杂 Word 版式设计
* 非中文排版场景

---

## 依赖

* python-docx >= 1.1.0 (别的版本没有测试)

---

## 未来计划

* [ ] 支持图片插入
* [ ] 更严格的 GB/T 9704 对齐

