# Vibe Docx V1.2 技术规格文档

> 本文档定义 V1.2 版本新增接口的详细实现规格，重点解决章节替换、Section 保护、精确插入等核心问题。

---

## 一、核心问题回顾

### 1.1 用户故事场景

用户执行"将 Markdown 内容合并到 Word 文档第5章"操作时遇到：

| 问题 | 根因 | 解决方案 |
|------|------|---------|
| 页边距变化 | 删除段落时误删 sectPr | Section 保护机制 |
| 图片丢失 | 无法智能定位图片位置 | 图片语义定位 |
| 只能追加 | 无精确插入接口 | 精确位置插入 |
| 样式不一致 | Markdown 转换硬编码样式 | 格式智能匹配 |

### 1.2 DOCX 结构理解

```
word/document.xml 结构：

<w:body>
  <w:p>...</w:p>                    <!-- 普通段落 -->
  <w:tbl>...</w:tbl>                <!-- 表格 -->
  <w:sectPr>...</w:sectPr>          <!-- Section 属性（可能在段落内或 body 末尾） -->
</w:body>

Section 边界标记：
1. <w:p><w:pPr><w:sectPr>...</w:sectPr></w:pPr></w:p>  <!-- 段落内 sectPr -->
2. <w:body>...<w:sectPr>...</w:sectPr></w:body>        <!-- body 末尾 sectPr -->

关键属性：
- w:pgSz (页面大小)
- w:pgMar (页边距)
- w:cols (分栏)
- w:docGrid (文档网格)
```

---

## 二、新增接口规格

### 2.1 Section 保护接口

#### detect_section_boundaries

```yaml
name: detect_section_boundaries
description: 检测文档中所有 section 边界位置
module: validator.py
version_added: "1.2"

input:
  file_path:
    type: string
    required: true
    description: DOCX 文件路径

output:
  success:
    type: boolean
    description: 操作是否成功
  sections:
    type: array
    description: Section 信息列表
    items:
      index:
        type: integer
        description: Section 索引（0-based）
      start_paragraph:
        type: integer
        description: 起始段落索引（包含）
      end_paragraph:
        type: integer
        description: 结束段落索引（sectPr 所在段落）
      sectPr_location:
        type: string
        enum: [in_paragraph, at_body_end]
        description: sectPr 位置类型
      page_settings:
        type: object
        properties:
          width:
            type: string
            description: 页面宽度 (EMU)
          height:
            type: string
            description: 页面高度 (EMU)
          margins:
            type: object
            properties:
              top: string
              bottom: string
              left: string
              right: string
          orientation:
            type: string
            enum: [portrait, landscape]
  total_count:
    type: integer
    description: Section 总数

errors:
  VAL001:
    message: 文档不存在
    recovery: 请提供有效的文档路径
  VAL003:
    message: 文档解析失败
    recovery: 检查文档是否损坏
```

**实现算法：**

```python
def detect_section_boundaries(file_path: str) -> Dict[str, Any]:
    """
    检测 section 边界。
    
    算法：
    1. 解压 DOCX，读取 word/document.xml
    2. 遍历 body 子元素
    3. 检测两种 sectPr 位置：
       - 段落内: <w:p><w:pPr><w:sectPr>
       - body 末尾: <w:body>...<w:sectPr>
    4. 提取页面设置属性
    """
    from docx import Document
    from docx.oxml.ns import qn
    
    doc = Document(file_path)
    body = doc._body._body
    
    sections = []
    current_start = 0
    
    for i, element in enumerate(body):
        # 检查段落内 sectPr
        if element.tag == qn('w:p'):
            sectPr = element.find('.//' + qn('w:sectPr'))
            if sectPr is not None:
                # 提取页面设置
                page_settings = _extract_page_settings(sectPr)
                sections.append({
                    "index": len(sections),
                    "start_paragraph": current_start,
                    "end_paragraph": i,  # sectPr 所在段落
                    "sectPr_location": "in_paragraph",
                    "page_settings": page_settings
                })
                current_start = i + 1
    
    # 检查 body 末尾 sectPr
    body_sectPr = body.find(qn('w:sectPr'))
    if body_sectPr is not None:
        page_settings = _extract_page_settings(body_sectPr)
        sections.append({
            "index": len(sections),
            "start_paragraph": current_start,
            "end_paragraph": len(doc.paragraphs),
            "sectPr_location": "at_body_end",
            "page_settings": page_settings
        })
    
    return {
        "success": True,
        "sections": sections,
        "total_count": len(sections)
    }
```

---

### 2.2 安全删除接口

#### delete_paragraphs

```yaml
name: delete_paragraphs
description: 安全删除段落范围，保护 section 结构
module: builder.py
version_added: "1.2"

input:
  session_id:
    type: string
    required: true
    description: 会话 ID
  start_index:
    type: integer
    required: true
    description: 起始段落索引（0-based）
  end_index:
    type: integer
    required: true
    description: 结束段落索引（不包含）
  options:
    type: object
    required: false
    properties:
      preserve_sections:
        type: boolean
        default: true
        description: 是否保护 section 设置
      preserve_images:
        type: boolean
        default: true
        description: 是否保留图片（移到删除范围外）
      preserve_tables:
        type: boolean
        default: false
        description: 是否保留表格

output:
  success:
    type: boolean
  deleted_count:
    type: integer
    description: 删除的段落数
  preserved_images:
    type: array
    description: 保留的图片信息
    items:
      original_paragraph: integer
      new_location: integer
      image_id: string
  section_preserved:
    type: boolean
    description: section 是否被保护

errors:
  BLD001:
    message: 会话不存在
    recovery: 请先调用 begin_session
  BLD004:
    message: 索引越界
    recovery: 检查段落索引范围
  BLD005:
    message: 删除范围包含 sectPr
    recovery: 使用 preserve_sections=true 或调整范围
```

**实现算法：**

```python
def delete_paragraphs(
    session_id: str,
    start_index: int,
    end_index: int,
    options: Optional[Dict] = None
) -> Dict[str, Any]:
    """
    安全删除段落。
    
    算法：
    1. 获取 section 边界信息
    2. 识别删除范围内的 sectPr
    3. 提取范围内的图片（如果 preserve_images）
    4. 删除段落元素
    5. 将 sectPr 移到安全位置
    6. 重新定位图片
    """
    opts = {
        "preserve_sections": True,
        "preserve_images": True,
        "preserve_tables": False
    }
    if options:
        opts.update(options)
    
    session = get_session(session_id)
    doc = Document(session.file_path)
    
    # 1. 获取 section 边界
    boundaries = detect_section_boundaries(session.file_path)
    sections = boundaries.get("sections", [])
    
    # 2. 识别范围内的 sectPr
    sectPrs_to_preserve = []
    for section in sections:
        if start_index <= section["end_paragraph"] < end_index:
            # sectPr 在删除范围内，需要保护
            sectPrs_to_preserve.append(section)
    
    # 3. 提取图片
    preserved_images = []
    if opts["preserve_images"]:
        for i in range(start_index, end_index):
            para = doc.paragraphs[i]
            drawings = para._element.findall('.//' + qn('w:drawing'))
            for drawing in drawings:
                # 记录图片信息
                preserved_images.append({
                    "original_paragraph": i,
                    "image_id": drawing.get(qn('r:id')),
                    "element": drawing
                })
    
    # 4. 删除段落
    deleted_count = 0
    elements_to_remove = []
    for i in range(start_index, end_index):
        if i < len(doc.paragraphs):
            elements_to_remove.append(doc.paragraphs[i]._element)
    
    for elem in elements_to_remove:
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)
            deleted_count += 1
    
    # 5. 保护 sectPr（移到删除范围前的段落）
    if opts["preserve_sections"] and sectPrs_to_preserve:
        for section in reversed(sectPrs_to_preserve):
            # 创建新的 sectPr 段落
            if start_index > 0:
                target_para = doc.paragraphs[start_index - 1]
                # 将 sectPr 附加到目标段落
                _attach_sectPr(target_para, section["page_settings"])
    
    doc.save(session.file_path)
    session.add_change({
        "type": "delete_paragraphs",
        "range": [start_index, end_index],
        "deleted_count": deleted_count
    })
    
    return {
        "success": True,
        "deleted_count": deleted_count,
        "preserved_images": preserved_images,
        "section_preserved": opts["preserve_sections"]
    }
```

---

### 2.3 精确插入接口

#### insert_after_paragraph

```yaml
name: insert_after_paragraph
description: 在指定段落后插入内容
module: builder.py
version_added: "1.2"

input:
  session_id:
    type: string
    required: true
  anchor:
    type: object
    required: true
    description: 锚点定位
    properties:
      type:
        type: string
        enum: [index, title, text]
        description: 锚点类型
      value:
        type: string | integer
        description: 锚点值（索引/标题/文本）
  content:
    type: object
    required: true
    description: 插入内容
    properties:
      type:
        type: string
        enum: [text, markdown, html]
        description: 内容类型
      text:
        type: string
        description: 文本内容
      style:
        type: object
        description: 样式选项
        properties:
          inherit:
            type: boolean
            default: true
            description: 是否继承周围样式
          font_name:
            type: string
          font_size:
            type: integer
          bold:
            type: boolean
  options:
    type: object
    properties:
      add_paragraph:
        type: boolean
        default: true
        description: 是否创建新段落

output:
  success:
    type: boolean
  inserted_at:
    type: integer
    description: 插入位置（新段落索引）
  style_applied:
    type: object
    description: 应用的样式信息

errors:
  BLD001:
    message: 会话不存在
  BLD006:
    message: 锚点未找到
    recovery: 检查锚点值是否正确
```

**实现算法：**

```python
def insert_after_paragraph(
    session_id: str,
    anchor: Dict,
    content: Dict,
    options: Optional[Dict] = None
) -> Dict[str, Any]:
    """
    精确位置插入。
    
    算法：
    1. 根据 anchor 类型定位锚点段落
    2. 解析内容（text/markdown/html）
    3. 确定插入位置
    4. 创建新段落元素
    5. 应用样式（继承或指定）
    6. 在 XML 层面插入
    """
    from docx.oxml import OxmlElement
    
    session = get_session(session_id)
    doc = Document(session.file_path)
    
    # 1. 定位锚点
    anchor_index = _find_anchor_index(doc, anchor)
    if anchor_index is None:
        return {
            "success": False,
            "error": {
                "code": "BLD006",
                "message": "锚点未找到",
                "recovery": f"检查 {anchor['type']}: {anchor['value']}"
            }
        }
    
    anchor_para = doc.paragraphs[anchor_index]
    
    # 2. 确定样式
    style_opts = content.get("style", {})
    if style_opts.get("inherit", True):
        # 继承锚点段落的样式
        inherited_style = _extract_paragraph_style(anchor_para)
    else:
        inherited_style = style_opts
    
    # 3. 创建新段落
    new_para = OxmlElement('w:p')
    
    # 4. 解析内容
    if content["type"] == "text":
        _add_text_to_paragraph(new_para, content["text"], inherited_style)
    elif content["type"] == "markdown":
        _add_markdown_to_paragraph(new_para, content["text"], inherited_style)
    
    # 5. 在锚点后插入
    anchor_elem = anchor_para._element
    anchor_elem.addnext(new_para)
    
    doc.save(session.file_path)
    
    return {
        "success": True,
        "inserted_at": anchor_index + 1,
        "style_applied": inherited_style
    }


def _find_anchor_index(doc, anchor: Dict) -> Optional[int]:
    """根据锚点类型查找段落索引"""
    anchor_type = anchor["type"]
    anchor_value = anchor["value"]
    
    if anchor_type == "index":
        return int(anchor_value)
    
    elif anchor_type == "title":
        for i, para in enumerate(doc.paragraphs):
            if para.style.name.startswith('Heading'):
                if anchor_value in para.text:
                    return i
    
    elif anchor_type == "text":
        for i, para in enumerate(doc.paragraphs):
            if anchor_value in para.text:
                return i
    
    return None


def _extract_paragraph_style(para) -> Dict:
    """提取段落样式"""
    style = {}
    
    # 字体
    for run in para.runs:
        if run.font.name:
            style["font_name"] = run.font.name
        if run.font.size:
            style["font_size"] = run.font.size.pt
        break
    
    # 段落属性
    pPr = para._element.find(qn('w:pPr'))
    if pPr is not None:
        # 行距
        spacing = pPr.find(qn('w:spacing'))
        if spacing is not None:
            line = spacing.get(qn('w:line'))
            if line:
                style["line_spacing"] = int(line) / 240  # 转换为倍数
        
        # 缩进
        indent = pPr.find(qn('w:ind'))
        if indent is not None:
            first_line = indent.get(qn('w:firstLine'))
            if first_line:
                style["first_line_indent"] = int(first_line) / 20  # 转换为磅
    
    return style
```

---

### 2.4 章节替换接口

#### replace_section

```yaml
name: replace_section
description: 替换章节内容，保留图片和 section 结构
module: builder.py
version_added: "1.2"

input:
  session_id:
    type: string
    required: true
  section_title:
    type: string
    required: true
    description: 要替换的章节标题（支持模糊匹配）
  new_content:
    type: object
    required: true
    description: 新内容
    properties:
      type:
        type: string
        enum: [markdown, text, paragraphs]
        description: 内容类型
      text:
        type: string
        description: 内容文本
  options:
    type: object
    properties:
      preserve_images:
        type: boolean
        default: true
        description: 保留原章节图片
      preserve_tables:
        type: boolean
        default: false
        description: 保留原章节表格
      image_position:
        type: string
        enum: [smart, end, start, preserve]
        default: smart
        description: 图片定位策略
      style_inherit:
        type: boolean
        default: true
        description: 继承文档样式
      style_template:
        type: string
        description: 样式模板名（可选）

output:
  success:
    type: boolean
  old_section:
    type: object
    description: 原章节信息
    properties:
      start_index: integer
      end_index: integer
      image_count: integer
      table_count: integer
  new_section:
    type: object
    description: 新章节信息
    properties:
      start_index: integer
      end_index: integer
      paragraph_count: integer
  images_relocated:
    type: array
    description: 重新定位的图片
  section_preserved:
    type: boolean
    description: section 设置是否保留

errors:
  BLD001:
    message: 会话不存在
  BLD007:
    message: 章节未找到
    recovery: 检查章节标题
  BLD008:
    message: 图片定位失败
    recovery: 使用 image_position=end 或手动指定
```

**实现算法：**

```python
def replace_section(
    session_id: str,
    section_title: str,
    new_content: Dict,
    options: Optional[Dict] = None
) -> Dict[str, Any]:
    """
    章节替换（保留图片和 section）。
    
    算法：
    1. 定位章节边界
    2. 提取章节内图片及其上下文
    3. 安全删除章节内容
    4. 解析新内容（Markdown → 段落）
    5. 插入新内容
    6. 智能定位图片
    7. 验证 section 完整性
    """
    opts = {
        "preserve_images": True,
        "preserve_tables": False,
        "image_position": "smart",
        "style_inherit": True
    }
    if options:
        opts.update(options)
    
    session = get_session(session_id)
    doc = Document(session.file_path)
    
    # 1. 定位章节
    section_info = _find_section(doc, section_title)
    if not section_info:
        return {
            "success": False,
            "error": {
                "code": "BLD007",
                "message": f"章节未找到: {section_title}"
            }
        }
    
    start_idx = section_info["start_index"]
    end_idx = section_info["end_index"]
    
    # 2. 提取图片及上下文
    images_with_context = []
    if opts["preserve_images"]:
        for i in range(start_idx, end_idx):
            para = doc.paragraphs[i]
            drawings = para._element.findall('.//' + qn('w:drawing'))
            for drawing in drawings:
                context = _extract_image_context(doc, i)
                images_with_context.append({
                    "element": drawing,
                    "original_index": i,
                    "context_before": context["before"],
                    "context_after": context["after"],
                    "caption": context["caption"]
                })
    
    old_image_count = len(images_with_context)
    
    # 3. 安全删除
    delete_result = delete_paragraphs(
        session_id,
        start_idx,
        end_idx,
        {
            "preserve_sections": True,
            "preserve_images": False,  # 已手动提取
            "preserve_tables": opts["preserve_tables"]
        }
    )
    
    # 4. 解析新内容
    new_paragraphs = []
    if new_content["type"] == "markdown":
        new_paragraphs = _parse_markdown(new_content["text"])
    elif new_content["type"] == "text":
        new_paragraphs = [{"type": "paragraph", "text": new_content["text"]}]
    
    # 5. 获取样式
    if opts["style_inherit"]:
        doc_style = get_body_style(session_id)
    else:
        doc_style = {}
    
    # 6. 插入新内容
    inserted_count = 0
    for i, para_data in enumerate(new_paragraphs):
        result = insert_after_paragraph(
            session_id,
            {"type": "index", "value": start_idx + i - 1 if i > 0 else start_idx - 1},
            {
                "type": "text",
                "text": para_data.get("text", ""),
                "style": doc_style
            }
        )
        if result["success"]:
            inserted_count += 1
    
    # 7. 智能定位图片
    images_relocated = []
    if opts["preserve_images"] and images_with_context:
        for img in images_with_context:
            new_position = _find_image_position(
                doc,
                img,
                start_idx,
                start_idx + inserted_count,
                opts["image_position"]
            )
            
            if new_position is not None:
                # 将图片插入到新位置
                _insert_image_at(doc, new_position, img["element"])
                images_relocated.append({
                    "original_paragraph": img["original_index"],
                    "new_location": new_position,
                    "method": "smart" if new_position != start_idx + inserted_count else "end"
                })
            else:
                # 无法定位，放到章节末尾
                _insert_image_at(doc, start_idx + inserted_count, img["element"])
                images_relocated.append({
                    "original_paragraph": img["original_index"],
                    "new_location": start_idx + inserted_count,
                    "method": "fallback"
                })
    
    doc.save(session.file_path)
    
    return {
        "success": True,
        "old_section": {
            "start_index": start_idx,
            "end_index": end_idx,
            "image_count": old_image_count,
            "table_count": 0
        },
        "new_section": {
            "start_index": start_idx,
            "end_index": start_idx + inserted_count,
            "paragraph_count": inserted_count
        },
        "images_relocated": images_relocated,
        "section_preserved": True
    }


def _extract_image_context(doc, para_index: int) -> Dict:
    """提取图片上下文"""
    context = {
        "before": "",
        "after": "",
        "caption": ""
    }
    
    # 前后各取 50 字符
    if para_index > 0:
        context["before"] = doc.paragraphs[para_index - 1].text[:50]
    if para_index < len(doc.paragraphs) - 1:
        context["after"] = doc.paragraphs[para_index + 1].text[:50]
    
    # 尝试识别图注（"图 X-X" 格式）
    para_text = doc.paragraphs[para_index].text
    caption_match = re.search(r'图\s*\d+-\d+[^\n]*', para_text)
    if caption_match:
        context["caption"] = caption_match.group()
    
    return context


def _find_image_position(doc, image_info: Dict, start: int, end: int, strategy: str) -> Optional[int]:
    """
    智能定位图片位置。
    
    策略：
    - smart: 根据上下文匹配最佳位置
    - end: 放到章节末尾
    - start: 放到章节开头
    - preserve: 尝试保持原位置比例
    """
    if strategy == "end":
        return end - 1
    elif strategy == "start":
        return start
    elif strategy == "smart":
        # 在新内容中搜索匹配的上下文
        for i in range(start, end):
            para_text = doc.paragraphs[i].text if i < len(doc.paragraphs) else ""
            
            # 检查是否有匹配的上下文
            if image_info["context_before"] and image_info["context_before"] in para_text:
                return i + 1
            if image_info["caption"] and image_info["caption"] in para_text:
                return i
        
        # 未找到匹配，返回 None（将使用 fallback）
        return None
    
    return end - 1
```

---

### 2.5 样式提取接口

#### get_body_style

```yaml
name: get_body_style
description: 提取文档正文样式
module: validator.py
version_added: "1.2"

input:
  session_id:
    type: string
    required: true

output:
  success:
    type: boolean
  body_style:
    type: object
    properties:
      font_name:
        type: string
        description: 正文字体
      font_size:
        type: integer
        description: 字号（磅）
      line_spacing:
        type: float
        description: 行距倍数
      first_line_indent:
        type: float
        description: 首行缩进（磅）
      paragraph_spacing:
        type: object
        properties:
          before: float
          after: float
  heading_styles:
    type: array
    description: 各级标题样式
    items:
      level: integer
      font_name: string
      font_size: integer
      bold: boolean

errors:
  BLD001:
    message: 会话不存在
```

**实现算法：**

```python
def get_body_style(session_id: str) -> Dict[str, Any]:
    """
    提取文档样式。
    
    策略：
    1. 统计正文段落样式频率
    2. 取最常见的样式作为 body_style
    3. 提取各级标题样式
    """
    session = get_session(session_id)
    doc = Document(session.file_path)
    
    # 统计正文样式
    font_counter = {}
    size_counter = {}
    line_spacing_sum = 0
    line_spacing_count = 0
    
    heading_styles = {}
    
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading'):
            # 标题样式
            level = 1
            if para.style.name != 'Heading':
                try:
                    level = int(para.style.name.split()[-1])
                except:
                    pass
            
            for run in para.runs:
                heading_styles[level] = {
                    "level": level,
                    "font_name": run.font.name or "黑体",
                    "font_size": run.font.size.pt if run.font.size else 14 + (2 - level) * 2,
                    "bold": run.bold or True
                }
                break
        else:
            # 正文样式
            for run in para.runs:
                if run.font.name:
                    font_counter[run.font.name] = font_counter.get(run.font.name, 0) + 1
                if run.font.size:
                    size_key = run.font.size.pt
                    size_counter[size_key] = size_counter.get(size_key, 0) + 1
                break
    
    # 提取行距
    for para in doc.paragraphs[:20]:  # 检查前20段
        pPr = para._element.find(qn('w:pPr'))
        if pPr is not None:
            spacing = pPr.find(qn('w:spacing'))
            if spacing is not None:
                line = spacing.get(qn('w:line'))
                if line:
                    line_spacing_sum += int(line) / 240
                    line_spacing_count += 1
    
    # 计算最常见样式
    body_style = {
        "font_name": max(font_counter.items(), key=lambda x: x[1])[0] if font_counter else "宋体",
        "font_size": max(size_counter.items(), key=lambda x: x[1])[0] if size_counter else 12,
        "line_spacing": line_spacing_sum / line_spacing_count if line_spacing_count > 0 else 1.5,
        "first_line_indent": 24,  # 默认 2 字符（约 24 磅）
        "paragraph_spacing": {
            "before": 0,
            "after": 0
        }
    }
    
    return {
        "success": True,
        "body_style": body_style,
        "heading_styles": [heading_styles.get(i) for i in sorted(heading_styles.keys())]
    }
```

---

## 三、数据结构定义

### 3.1 SectionInfo

```python
@dataclass
class SectionInfo:
    """Section 信息"""
    index: int                          # Section 索引
    start_paragraph: int                # 起始段落
    end_paragraph: int                  # 结束段落
    sectPr_location: str                # "in_paragraph" | "at_body_end"
    page_settings: PageSettings          # 页面设置
    
@dataclass
class PageSettings:
    """页面设置"""
    width: int                          # 页面宽度 (EMU)
    height: int                         # 页面高度 (EMU)
    margins: Margins                    # 页边距
    orientation: str                    # "portrait" | "landscape"
    
@dataclass
class Margins:
    """页边距"""
    top: int                            # EMU
    bottom: int
    left: int
    right: int
```

### 3.2 ImageContext

```python
@dataclass
class ImageContext:
    """图片上下文"""
    image_id: str                       # 图片 ID (rId)
    original_paragraph: int             # 原段落索引
    context_before: str                 # 前文上下文
    context_after: str                  # 后文上下文
    caption: str                        # 图注（如有）
    element: Any                        # XML 元素引用
```

### 3.3 InsertAnchor

```python
class AnchorType(Enum):
    INDEX = "index"                     # 段落索引
    TITLE = "title"                     # 标题文本
    TEXT = "text"                       # 文本匹配

@dataclass
class InsertAnchor:
    """插入锚点"""
    type: AnchorType
    value: Union[int, str]
    
    def find_index(self, doc: Document) -> Optional[int]:
        """查找锚点对应的段落索引"""
        # 实现见上文
        pass
```

---

## 四、错误处理

### 4.1 错误码定义

```python
class BuilderError(Exception):
    """Builder 错误"""
    
    ERROR_CODES = {
        "BLD001": {
            "message": "会话不存在",
            "recovery": "请先调用 begin_session",
            "can_retry": False
        },
        "BLD002": {
            "message": "操作执行失败",
            "recovery": "自动回滚，恢复原文件",
            "can_retry": True
        },
        "BLD003": {
            "message": "操作后验证失败",
            "recovery": "询问用户是否保留修改",
            "can_retry": False
        },
        "BLD004": {
            "message": "索引越界",
            "recovery": "检查段落索引范围",
            "can_retry": False
        },
        "BLD005": {
            "message": "删除范围包含 sectPr",
            "recovery": "使用 preserve_sections=true 或调整范围",
            "can_retry": True
        },
        "BLD006": {
            "message": "锚点未找到",
            "recovery": "检查锚点值是否正确",
            "can_retry": False
        },
        "BLD007": {
            "message": "章节未找到",
            "recovery": "检查章节标题",
            "can_retry": False
        },
        "BLD008": {
            "message": "图片定位失败",
            "recovery": "使用 image_position=end 或手动指定",
            "can_retry": True
        }
    }
```

### 4.2 统一响应格式

```python
def success_response(data: Dict, stats: Optional[Dict] = None) -> Dict:
    """成功响应"""
    return {
        "success": True,
        "data": data,
        "stats": stats or {}
    }

def error_response(code: str, detail: str = "") -> Dict:
    """错误响应"""
    info = BuilderError.ERROR_CODES.get(code, {})
    return {
        "success": False,
        "error": {
            "code": code,
            "message": info.get("message", "未知错误"),
            "detail": detail,
            "recovery": info.get("recovery", ""),
            "can_retry": info.get("can_retry", False)
        }
    }
```

---

## 五、测试用例

### 5.1 Section 保护测试

```python
def test_detect_section_boundaries():
    """测试 section 边界检测"""
    # 简单文档（1 section）
    result = detect_section_boundaries("tests/fixtures/simple.docx")
    assert result["success"]
    assert result["total_count"] == 1
    
    # 混合页面（横向+纵向）
    result = detect_section_boundaries("tests/fixtures/mixed_orientation.docx")
    assert result["success"]
    assert result["total_count"] >= 2
    
    # 验证页面设置提取
    for section in result["sections"]:
        assert "page_settings" in section
        assert "margins" in section["page_settings"]

def test_delete_preserves_sections():
    """测试删除不破坏 section"""
    session = begin_session("tests/fixtures/mixed_sections.docx")
    
    # 记录原 section 数量
    before = detect_section_boundaries(session["file_path"])
    
    # 删除段落
    delete_paragraphs(session["session_id"], 10, 20, {"preserve_sections": True})
    commit(session["session_id"])
    
    # 验证 section 数量不变
    after = detect_section_boundaries(session["file_path"])
    assert after["total_count"] == before["total_count"]
    
    # 验证页边距不变
    for i, section in enumerate(after["sections"]):
        assert section["page_settings"]["margins"] == before["sections"][i]["page_settings"]["margins"]
```

### 5.2 精确插入测试

```python
def test_insert_after_index():
    """测试索引锚点插入"""
    session = begin_session("tests/fixtures/simple.docx")
    
    result = insert_after_paragraph(
        session["session_id"],
        {"type": "index", "value": 5},
        {"type": "text", "text": "插入的段落"}
    )
    
    assert result["success"]
    assert result["inserted_at"] == 6

def test_insert_after_title():
    """测试标题锚点插入"""
    session = begin_session("tests/fixtures/report.docx")
    
    result = insert_after_paragraph(
        session["session_id"],
        {"type": "title", "value": "第三章"},
        {"type": "text", "text": "新内容"}
    )
    
    assert result["success"]
    # 验证插入位置正确

def test_insert_inherits_style():
    """测试样式继承"""
    session = begin_session("tests/fixtures/styled.docx")
    
    result = insert_after_paragraph(
        session["session_id"],
        {"type": "index", "value": 0},
        {
            "type": "text",
            "text": "继承样式的段落",
            "style": {"inherit": True}
        }
    )
    
    assert result["success"]
    assert "font_name" in result["style_applied"]
```

### 5.3 章节替换测试

```python
def test_replace_section_preserves_images():
    """测试章节替换保留图片"""
    session = begin_session("tests/fixtures/chapter_with_images.docx")
    
    # 原图片数量
    doc = Document(session["file_path"])
    original_images = count_images(doc)
    
    result = replace_section(
        session["session_id"],
        "第五章",
        {"type": "markdown", "text": "# 新内容\n\n这是新的内容。"},
        {"preserve_images": True, "image_position": "smart"}
    )
    
    assert result["success"]
    assert len(result["images_relocated"]) > 0
    
    # 验证图片数量不变
    doc_after = Document(session["file_path"])
    assert count_images(doc_after) == original_images

def test_replace_section_preserves_margins():
    """测试章节替换保留页边距"""
    session = begin_session("tests/fixtures/report.docx")
    
    # 原页边距
    before = detect_section_boundaries(session["file_path"])
    original_margins = before["sections"][-1]["page_settings"]["margins"]
    
    result = replace_section(
        session["session_id"],
        "第五章",
        {"type": "text", "text": "新内容"},
        {"preserve_images": False}
    )
    
    # 验证页边距不变
    after = detect_section_boundaries(session["file_path"])
    assert after["sections"][-1]["page_settings"]["margins"] == original_margins
```

---

## 六、实现优先级

| 优先级 | 接口 | 预计工作量 | 依赖 |
|--------|------|-----------|------|
| P0-1 | `detect_section_boundaries` | 2h | 无 |
| P0-2 | `delete_paragraphs` | 2h | P0-1 |
| P0-3 | `insert_after_paragraph` | 2h | 无 |
| P0-4 | `get_body_style` | 1h | 无 |
| P0-5 | `replace_section` | 3h | P0-1, P0-2, P0-3, P0-4 |

**总工作量：约 10 小时**

---

## 七、验收标准

### 7.1 功能验收

- [ ] `detect_section_boundaries` 正确识别所有 section 边界
- [ ] `delete_paragraphs` 删除后 section 数量不变
- [ ] `delete_paragraphs` 删除后页边距不变
- [ ] `insert_after_paragraph` 插入位置准确
- [ ] `insert_after_paragraph` 样式继承正确
- [ ] `replace_section` 图片数量不变
- [ ] `replace_section` 图片位置合理

### 7.2 质量验收

- [ ] 单元测试覆盖率 ≥ 80%
- [ ] 集成测试通过率 100%
- [ ] 无严重 bug
- [ ] API 文档完整
