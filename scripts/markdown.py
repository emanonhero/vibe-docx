# -*- coding: utf-8 -*-
"""
Markdown 解析器

将 Markdown 格式转换为 Word XML 格式。
"""

import re
from typing import List, Tuple


def escape_xml(text: str) -> str:
    """转义 XML 特殊字符"""
    return (text
        .replace('&', '&amp;')
        .replace('<', '&lt;')
        .replace('>', '&gt;')
        .replace('"', '&quot;')
        .replace("'", '&apos;'))


class MarkdownParser:
    """
    Markdown 到 Word XML 转换器。
    
    支持的 Markdown 格式：
    - 标题（# ~ ######）
    - 加粗（**text** 或 __text__）
    - 斜体（*text* 或 _text_）
    - 删除线（~~text~~）
    - 无序列表（- 或 *）
    - 有序列表（1. 2. 3.）
    - 表格
    - 图片占位符 {{image:rId}}
    """
    
    def __init__(self):
        """初始化解析器"""
        self.heading_styles = {
            1: 'Heading1',
            2: 'Heading2',
            3: 'Heading3',
            4: 'Heading4',
            5: 'Heading5',
            6: 'Heading6',
        }
    
    def parse(self, content: str) -> str:
        """
        解析 Markdown 内容并生成 Word XML。
        
        Args:
            content: Markdown 格式的内容
            
        Returns:
            Word XML 字符串
        """
        xml_parts = []
        lines = content.split('\n')
        i = 0
        
        while i < len(lines):
            line = lines[i].strip()
            
            # 空行
            if not line:
                i += 1
                continue
            
            # 图片占位符
            img_match = re.match(r'^\{\{image:(rId\d+)\}\}', line)
            if img_match:
                rid = img_match.group(1)
                xml_parts.append(self._generate_image_placeholder(rid))
                i += 1
                continue
            
            # 标题
            heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if heading_match:
                level = len(heading_match.group(1))
                text = heading_match.group(2)
                xml_parts.append(self._generate_heading(text, level))
                i += 1
                continue
            
            # 无序列表
            if re.match(r'^[-*]\s+', line):
                list_items, i = self._parse_unordered_list(lines, i)
                xml_parts.extend(list_items)
                continue
            
            # 有序列表
            if re.match(r'^\d+\.\s+', line):
                list_items, i = self._parse_ordered_list(lines, i)
                xml_parts.extend(list_items)
                continue
            
            # 表格
            if line.startswith('|'):
                table_xml, i = self._parse_table(lines, i)
                xml_parts.append(table_xml)
                continue
            
            # 水平线
            if re.match(r'^[-*_]{3,}$', line):
                xml_parts.append(self._generate_horizontal_line())
                i += 1
                continue
            
            # 普通段落
            xml_parts.append(self._generate_paragraph(line))
            i += 1
        
        return '\n'.join(xml_parts)
    
    def _generate_heading(self, text: str, level: int) -> str:
        """生成标题 XML"""
        style = self.heading_styles.get(level, 'Heading1')
        runs = self._parse_inline_formatting(text)
        return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{runs}</w:p>'
    
    def _generate_paragraph(self, text: str) -> str:
        """生成段落 XML"""
        runs = self._parse_inline_formatting(text)
        return f'<w:p>{runs}</w:p>'
    
    def _parse_inline_formatting(self, text: str) -> str:
        """解析行内格式（加粗、斜体、删除线）"""
        if not text:
            return '<w:r><w:t/></w:r>'
        
        runs = []
        pos = 0
        
        # 匹配格式：***加粗斜体*** | **加粗** | *斜体* | ~~删除线~~
        pattern = r'\*\*\*(.+?)\*\*\*|(\*\*|__)(.+?)\2|(\*|_)(.+?)\4|~~(.+?)~~'
        
        for match in re.finditer(pattern, text):
            # 添加前面的普通文本
            if match.start() > pos:
                plain_text = text[pos:match.start()]
                runs.append(self._generate_run(plain_text))
            
            # 添加格式化文本
            if match.group(1):  # ***加粗斜体***
                run_text = escape_xml(match.group(1))
                runs.append(f'<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>{run_text}</w:t></w:r>')
            elif match.group(2):  # **加粗** 或 __加粗__
                run_text = escape_xml(match.group(3))
                runs.append(f'<w:r><w:rPr><w:b/></w:rPr><w:t>{run_text}</w:t></w:r>')
            elif match.group(4):  # *斜体* 或 _斜体_
                run_text = escape_xml(match.group(5))
                runs.append(f'<w:r><w:rPr><w:i/></w:rPr><w:t>{run_text}</w:t></w:r>')
            elif match.group(6):  # ~~删除线~~
                run_text = escape_xml(match.group(6))
                runs.append(f'<w:r><w:rPr><w:strike/></w:rPr><w:t>{run_text}</w:t></w:r>')
            
            pos = match.end()
        
        # 添加剩余文本
        if pos < len(text):
            runs.append(self._generate_run(text[pos:]))
        
        if not runs:
            runs.append('<w:r><w:t/></w:r>')
        
        return ''.join(runs)
    
    def _generate_run(self, text: str) -> str:
        """生成单个 run XML"""
        escaped_text = escape_xml(text)
        return f'<w:r><w:t>{escaped_text}</w:t></w:r>'
    
    def _parse_unordered_list(self, lines: List[str], start_idx: int) -> Tuple[List[str], int]:
        """解析无序列表"""
        items = []
        i = start_idx
        
        while i < len(lines):
            line = lines[i].strip()
            match = re.match(r'^[-*]\s+(.+)$', line)
            if match:
                text = match.group(1)
                runs = self._parse_inline_formatting(text)
                items.append(f'<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr>{runs}</w:p>')
                i += 1
            else:
                break
        
        return items, i
    
    def _parse_ordered_list(self, lines: List[str], start_idx: int) -> Tuple[List[str], int]:
        """解析有序列表"""
        items = []
        i = start_idx
        
        while i < len(lines):
            line = lines[i].strip()
            match = re.match(r'^\d+\.\s+(.+)$', line)
            if match:
                text = match.group(1)
                runs = self._parse_inline_formatting(text)
                items.append(f'<w:p><w:pPr><w:pStyle w:val="ListNumber"/></w:pPr>{runs}</w:p>')
                i += 1
            else:
                break
        
        return items, i
    
    def _parse_table(self, lines: List[str], start_idx: int) -> Tuple[str, int]:
        """解析表格"""
        rows = []
        i = start_idx
        
        while i < len(lines):
            line = lines[i].strip()
            if line.startswith('|'):
                # 跳过分隔行
                if re.match(r'^\|[\s\-:|]+\|$', line):
                    i += 1
                    continue
                
                cells = [c.strip() for c in line.split('|') if c.strip()]
                cell_xml = ''.join(
                    f'<w:tc><w:p>{self._parse_inline_formatting(c)}</w:p></w:tc>'
                    for c in cells
                )
                rows.append(f'<w:tr>{cell_xml}</w:tr>')
                i += 1
            else:
                break
        
        # 生成表格属性（包括边框）
        tbl_pr = '''<w:tblPr>
<w:tblW w:w="0" w:type="auto"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders>
</w:tblPr>'''
        
        table_xml = f'<w:tbl>{tbl_pr}{"".join(rows)}</w:tbl>'
        return table_xml, i
    
    def _generate_image_placeholder(self, rid: str) -> str:
        """生成图片占位符 XML"""
        bookmark_name = f"IMAGE_PLACEHOLDER_{rid}"
        return f'''<w:p>
    <w:bookmarkStart w:id="0" w:name="{bookmark_name}"/>
    <w:bookmarkEnd w:id="0"/>
    <w:r><w:t>[图片占位符: {rid}]</w:t></w:r>
</w:p>'''
    
    def _generate_horizontal_line(self) -> str:
        """生成水平线 XML"""
        return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>'


def parse_markdown_to_xml(content: str) -> str:
    """
    将 Markdown 内容转换为 Word XML。
    
    Args:
        content: Markdown 格式的内容
        
    Returns:
        Word XML 字符串
        
    Example:
        >>> xml = parse_markdown_to_xml('# 标题\n这是**加粗**文本')
    """
    parser = MarkdownParser()
    return parser.parse(content)
