# 模板定义

本文档说明预置模板的配置和自定义方法。

## 预置模板

### business_report - 商务报告

适合商业提案、项目报告、正式文档。

```python
{
    "title_styles": {
        "Heading 1": {"font_name": "黑体", "font_size": 18, "bold": True},
        "Heading 2": {"font_name": "黑体", "font_size": 14, "bold": True}
    },
    "paragraph_style": {
        "line_spacing": 1.5,
        "font_name": "宋体",
        "font_size": 12
    },
    "page_setup": {
        "margins": {
            "top": "2.54cm",
            "bottom": "2.54cm",
            "left": "3.17cm",
            "right": "3.17cm"
        }
    }
}
```

### internal_simple - 简洁风格

适合内部通知、简洁记录。

```python
{
    "title_styles": {
        "Title": {"font_name": "黑体", "font_size": 16, "bold": True}
    },
    "paragraph_style": {
        "line_spacing": 1.25,
        "font_name": "宋体",
        "font_size": 12
    }
}
```

### academic_paper - 学术论文

适合研究报告、学术文档。

```python
{
    "title_styles": {
        "Heading 1": {"font_name": "黑体", "font_size": 16, "bold": True},
        "Heading 2": {"font_name": "黑体", "font_size": 14, "bold": True}
    },
    "paragraph_style": {
        "line_spacing": 1.5,
        "first_line_indent": "2ch",  # 首行缩进
        "font_name": "宋体",
        "font_size": 12
    },
    "page_setup": {
        "margins": {
            "top": "2.5cm",
            "bottom": "2.5cm",
            "left": "3cm",
            "right": "3cm"
        }
    }
}
```

## 使用模板

### 获取模板配置

```python
from scripts.builder import get_template

template = get_template("business_report")
if template["success"]:
    config = template["template"]
    print(config["paragraph_style"])
```

### 应用模板

```python
from scripts.builder import begin_session, fix_page_setup, fix_formatting, commit

session = begin_session("report.docx", backup=True)
session_id = session["session_id"]

# 应用页面设置
template = get_template("business_report")["template"]
page_setup = template.get("page_setup", {})

if page_setup.get("margins"):
    fix_page_setup(session_id, margins=page_setup["margins"])

# 应用格式
fix_formatting(session_id, options={
    "default_font": template["paragraph_style"]["font_name"],
    "title_font": template["title_styles"]["Heading 1"]["font_name"]
})

commit(session_id)
```

## 自定义模板

可以通过代码定义自定义模板：

```python
# 定义自定义模板
custom_template = {
    "title_styles": {
        "Heading 1": {"font_name": "微软雅黑", "font_size": 20, "bold": True, "color": "2E74B5"},
        "Heading 2": {"font_name": "微软雅黑", "font_size": 16, "bold": True}
    },
    "paragraph_style": {
        "line_spacing": 1.5,
        "font_name": "微软雅黑",
        "font_size": 11,
        "space_after": "6pt"
    },
    "page_setup": {
        "margins": {
            "top": "2cm",
            "bottom": "2cm",
            "left": "2.5cm",
            "right": "2.5cm"
        }
    }
}

# 应用自定义模板
session = begin_session("report.docx", backup=True)
session_id = session["session_id"]

# 应用页面设置
fix_page_setup(session_id, margins=custom_template["page_setup"]["margins"])

# 应用格式
fix_formatting(session_id, options={
    "default_font": custom_template["paragraph_style"]["font_name"],
    "title_font": custom_template["title_styles"]["Heading 1"]["font_name"]
})

commit(session_id)
```

## 模板参数说明

### title_styles

标题样式配置，key 为 Word 样式名称。

| 参数 | 类型 | 说明 |
|------|------|------|
| `font_name` | str | 字体名称 |
| `font_size` | int | 字号（磅） |
| `bold` | bool | 是否加粗 |
| `color` | str | 颜色（十六进制，可选） |

### paragraph_style

段落样式配置。

| 参数 | 类型 | 说明 |
|------|------|------|
| `font_name` | str | 字体名称 |
| `font_size` | int | 字号（磅） |
| `line_spacing` | float | 行距倍数 |
| `first_line_indent` | str | 首行缩进，如 "2ch" |
| `space_after` | str | 段后间距，如 "6pt" |

### page_setup

页面设置。

| 参数 | 类型 | 说明 |
|------|------|------|
| `margins` | dict | 页边距 |
| `margins.top` | str | 上边距，如 "2.54cm" |
| `margins.bottom` | str | 下边距 |
| `margins.left` | str | 左边距 |
| `margins.right` | str | 右边距 |

## 常用字体

| 中文名 | 英文名 | 适用场景 |
|-------|-------|---------|
| 宋体 | SimSun | 正文 |
| 黑体 | SimHei | 标题 |
| 微软雅黑 | Microsoft YaHei | 现代、商务 |
| 楷体 | KaiTi | 公文、正式 |
| 仿宋 | FangSong | 公文 |

## 页面尺寸

常用纸张尺寸（用于 `fix_page_setup` 的 `page_size` 参数）：

```python
# A4
page_size = {"width": "21cm", "height": "29.7cm"}

# A3
page_size = {"width": "29.7cm", "height": "42cm"}

# Letter
page_size = {"width": "8.5in", "height": "11in"}

# 16K
page_size = {"width": "18.4cm", "height": "26cm"}
```
