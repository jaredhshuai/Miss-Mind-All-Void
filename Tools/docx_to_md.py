#!/usr/bin/env python3
"""将 Word 文档转换为 Markdown 格式，并自动修复格式问题

功能特性：
1. 保留文档顺序（段落、表格、列表、标题）
2. 导出图片到 Images/<basename>_images/ 并使用相对路径引用
3. 支持文件名规范化（--normalize 选项）
4. 处理手动换行符 <w:br/> 和控制字符
5. 合并相同样式的连续文本块，减少重复的 Markdown 标记
6. 修复中文粗体引号格式问题（引号移到粗体标记外）
7. 统一标点符号编码为标准中文标点
8. 后处理清理：
   - 替换 NBSP 为普通空格
   - 分离首行完整粗体句子为独立段落
   - 合并相邻的粗体段落
   - 将中文标点符号移到粗体块外
   - 为非空行添加尾部两空格（软换行）
9. 幂等性：重复运行产生稳定的 Markdown 输出

使用方法：
    python docx_to_md.py <docx_path> [--normalize]
    
    --normalize: 规范化文件名（移除特殊字符，用连字符替换空格）
"""

import sys
import re
from pathlib import Path
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn


# ============================================================================
# 工具函数：文件名规范化
# ============================================================================

def normalize_name(name: str) -> str:
    """规范化文件名：移除特殊标点，替换空格为连字符"""
    # 移除可能导致渲染或路径问题的特殊标点
    name = re.sub(r'[《》："""\:]', '', name)
    # 将空白字符序列替换为单个连字符
    name = re.sub(r'\s+', '-', name.strip())
    # 合并多个连字符
    name = re.sub(r'-{2,}', '-', name)
    return name.strip('-')


# ============================================================================
# 工具函数：标点符号规范化
# ============================================================================

def normalize_punctuation(text: str) -> str:
    """统一标点符号编码为标准中文标点
    
    采用配对策略处理引号，避免破坏代码块等内容
    """
    # 统一双引号：使用配对策略，奇数次用左引号，偶数次用右引号
    quote_count = 0
    result = []
    for char in text:
        if char in ['\u0022', '\uff02']:  # ASCII " 或全角 ＂
            if quote_count % 2 == 0:
                result.append('\u201c')  # 左引号 "
            else:
                result.append('\u201d')  # 右引号 "
            quote_count += 1
        else:
            result.append(char)
    text = ''.join(result)
    
    # 统一单引号：只在中文语境中替换
    single_quote_count = 0
    result = []
    for i, char in enumerate(text):
        if char == '\u0027':  # ASCII单引号 '
            # 检查前后是否有中文字符
            prev_is_chinese = i > 0 and '\u4e00' <= text[i-1] <= '\u9fff'
            next_is_chinese = i < len(text) - 1 and '\u4e00' <= text[i+1] <= '\u9fff'
            
            if prev_is_chinese or next_is_chinese:
                if single_quote_count % 2 == 0:
                    result.append('\u2018')  # 左单引号 '
                else:
                    result.append('\u2019')  # 右单引号 '
                single_quote_count += 1
            else:
                result.append(char)
        else:
            result.append(char)
    text = ''.join(result)
    
    # 统一其他标点：仅在中文语境中替换
    # 感叹号
    text = re.sub(r'([\u4e00-\u9fff])!', r'\1！', text)
    text = re.sub(r'!([\u4e00-\u9fff])', r'！\1', text)
    
    # 问号
    text = re.sub(r'([\u4e00-\u9fff])\?', r'\1？', text)
    text = re.sub(r'\?([\u4e00-\u9fff])', r'？\1', text)
    
    # 逗号
    text = re.sub(r'([\u4e00-\u9fff]),([\u4e00-\u9fff])', r'\1，\2', text)
    
    # 分号
    text = re.sub(r'([\u4e00-\u9fff]);', r'\1；', text)
    text = re.sub(r';([\u4e00-\u9fff])', r'；\1', text)
    
    # 冒号
    text = re.sub(r'([\u4e00-\u9fff]):', r'\1：', text)
    text = re.sub(r':([\u4e00-\u9fff])', r'：\1', text)
    
    # 括号
    text = re.sub(r'([\u4e00-\u9fff])\(', r'\1（', text)
    text = re.sub(r'\(([\u4e00-\u9fff])', r'（\1', text)
    text = re.sub(r'([\u4e00-\u9fff])\)', r'\1）', text)
    text = re.sub(r'\)([\u4e00-\u9fff])', r'）\1', text)
    
    return text


def fix_bold_quotes(text: str) -> str:
    """修复粗体引号格式：将引号从粗体标记内移到外部
    
    转换规则：
    - **"文本"** -> **"文本"**
    - "**文本**" -> **"文本"**
    - "**文本**。" -> **"文本。"**
    - **"xxx**" -> **"xxx"**
    """
    # 第一步：统一所有引号编码为标准中文引号
    text = normalize_punctuation(text)
    
    # 第二步：进行格式转换（现在只需要处理统一后的中文引号）
    
    # 规则0：**"xxx"yyy** -> **"xxx"**yyy（处理引号在粗体开头的情况）
    text = re.sub(r'\*\*"([^"]*?)"([^*]*?)\*\*', r'**"\1"**\2', text)
    
    # 规则1："**文本**" -> **"文本"**
    # 匹配：开引号 + 粗体 + 结束引号
    text = re.sub(r'"(\*\*([^*]+?)\*\*)"', lambda m: f'**"{m.group(2)}"**', text)
    
    # 规则2："**文本**。" -> **"文本。"**
    # 匹配：开引号 + 粗体 + 标点 + 结束引号
    text = re.sub(r'"(\*\*([^*]+?)\*\*)([。，、！？；：])"', lambda m: f'**"{m.group(2)}{m.group(3)}"**', text)
    
    # 规则3：**"xxx**" -> **"xxx"**
    # 处理粗体开始正确但结束位置错误的情况
    text = re.sub(r'\*\*"([^"]+?)\*\*"', r'**"\1"**', text)
    
    return text


# ============================================================================
# 文本提取与样式处理
# ============================================================================

def _extract_run_style(run):
    """提取单个 run 的文本与样式，不直接返回已包裹的 markdown，方便后续合并"""
    return {
        'text': run.text or '',
        'bold': bool(run.bold),
        'italic': bool(run.italic),
        'is_image': False,
        'is_line_break': False,
    }


_EXTRA_BREAK_CHARS_PATTERN = re.compile(r'[\r\n\u000b\u000c\u2028\u2029]')


def _split_text_with_breaks(text: str) -> list:
    """将包含额外换行控制符的文本拆分为纯文本和显式换行 token
    
    使用与手动 <w:br/> 相同的 '  \\n' 作为换行标记，后续逻辑即可复用
    """
    if not text:
        return []
    parts = []
    last = 0
    for m in _EXTRA_BREAK_CHARS_PATTERN.finditer(text):
        seg = text[last:m.start()]
        if seg:
            parts.append({'text': seg, 'is_line_break': False})
        parts.append({'text': '  \n', 'is_line_break': True})
        last = m.end()
    tail = text[last:]
    if tail:
        parts.append({'text': tail, 'is_line_break': False})
    return parts


def _format_merged(text: str, bold: bool, italic: bool) -> str:
    """将文本格式化为 Markdown 样式"""
    if not text:
        return ''
    if bold and italic:
        return f"***{text}***"
    if bold:
        return f"**{text}**"
    if italic:
        return f"*{text}*"
    return text


def _iter_block_items(doc):
    """按文档顺序生成段落和表格"""
    for child in doc.element.body:
        if child.tag.endswith('p'):
            yield Paragraph(child, doc)
        elif child.tag.endswith('tbl'):
            yield Table(child, doc)


def _runs_to_tokens(paragraph, part, images_dir: Path, exported: set) -> list:
    """将段落的 runs 转换为 Markdown tokens"""
    # 采集结构化 run
    raw_runs = []
    for run in getattr(paragraph, 'runs', []):
        data = _extract_run_style(run)
        inline = run.element
        
        # 处理内部手动换行 <w:br/>
        if inline.xpath('.//w:br'):
            raw_runs.append({
                'text': '  \n',
                'bold': False,
                'italic': False,
                'is_image': False,
                'is_line_break': True
            })
        
        # 处理图像
        blips = inline.xpath('.//a:blip')
        for blip in blips:
            rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if not rId:
                continue
            image_part = part.related_parts.get(rId)
            if image_part is None:
                continue
            
            # 使用 rId 作为唯一标识，避免重名文件覆盖
            base_filename = Path(image_part.filename).name
            if rId in exported:
                # 已经导出过这个 rId，直接使用之前的文件名
                actual_filename = exported[rId]
            else:
                # 检查文件名是否重复，如果重复则添加序号
                actual_filename = base_filename
                counter = 1
                while (images_dir / actual_filename).exists():
                    stem = Path(base_filename).stem
                    suffix = Path(base_filename).suffix
                    actual_filename = f"{stem}_{counter}{suffix}"
                    counter += 1
                
                # 导出图片
                out_path = images_dir / actual_filename
                with open(out_path, 'wb') as f:
                    f.write(image_part.blob)
                exported[rId] = actual_filename
            
            raw_runs.append({
                'text': f"![](Images/{images_dir.name}/{actual_filename})",
                'bold': False,
                'italic': False,
                'is_image': True,
                'is_line_break': False
            })
        
        # 拆分 run.text 中的额外换行控制符
        txt = data['text']
        if txt:
            if _EXTRA_BREAK_CHARS_PATTERN.search(txt):
                for part_item in _split_text_with_breaks(txt):
                    raw_runs.append({
                        'text': part_item['text'],
                        'bold': data['bold'],
                        'italic': data['italic'],
                        'is_image': False,
                        'is_line_break': part_item['is_line_break'],
                    })
            else:
                raw_runs.append(data)
    
    # Fallback：如果 runs 为空但段落有文本，直接使用 paragraph.text
    if not raw_runs and paragraph.text.strip():
        raw_runs.append({
            'text': paragraph.text,
            'bold': False,
            'italic': False,
            'is_image': False,
            'is_line_break': False,
        })

    # 合并连续样式一致（且非图片/换行）的 run，减少重复 ** 包裹
    merged = []
    for item in raw_runs:
        if not merged:
            merged.append(item)
            continue
        prev = merged[-1]
        if (not item['is_image'] and not item['is_line_break'] and
            not prev['is_image'] and not prev['is_line_break'] and
            item['bold'] == prev['bold'] and item['italic'] == prev['italic']):
            prev['text'] += item['text']
        else:
            merged.append(item)

    # 输出为最终 markdown token 列表
    tokens = []
    for i, m in enumerate(merged):
        if m['is_image']:
            tokens.append(m['text'])
            # 如果图片后还有内容（不是最后一项），添加换行
            if i < len(merged) - 1:
                tokens.append('\n\n')
        elif m['is_line_break']:
            tokens.append(m['text'])
        else:
            formatted = _format_merged(m['text'], m['bold'], m['italic'])
            if formatted:
                tokens.append(formatted)
    return tokens


# ============================================================================
# 列表处理
# ============================================================================

def _is_list_paragraph(paragraph) -> bool:
    """检查段落是否为列表项"""
    return bool(paragraph._p.xpath('.//w:numPr'))


def _list_marker(paragraph) -> str:
    """获取列表标记"""
    numFmt = paragraph._p.xpath('.//w:numPr//w:numFmt')
    if numFmt:
        val = numFmt[0].get(qn('w:val'))
        if val == 'bullet':
            return '- '
        return '1. '
    return '- '


def _get_list_number(paragraph, counters: dict, part) -> int | None:
    """获取有序列表的实际编号"""
    numPr = paragraph._p.xpath('.//w:numPr')
    if not numPr:
        return None
    numId_e = numPr[0].xpath('./w:numId')
    ilvl_e = numPr[0].xpath('./w:ilvl')
    numFmt_e = numPr[0].xpath('./w:numFmt')
    if not numId_e or not ilvl_e or not numFmt_e:
        return None
    if numFmt_e[0].get(qn('w:val')) == 'bullet':
        return None
    numId = numId_e[0].get(qn('w:val'))
    ilvl = ilvl_e[0].get(qn('w:val'))
    key = (numId, ilvl)
    numbering = part.numbering_part.element
    abstract_e = numbering.xpath(f'.//w:num[w:numId[@w:val="{numId}"]]/w:abstractNumId')
    start_val = 1
    if abstract_e:
        abstract_id = abstract_e[0].get(qn('w:val'))
        start_e = numbering.xpath(
            f'.//w:abstractNum[w:abstractNumId[@w:val="{abstract_id}"]]'
            f'//w:lvl[@w:ilvl="{ilvl}"]//w:start'
        )
        if start_e:
            try:
                start_val = int(start_e[0].get(qn('w:val')))
            except ValueError:
                start_val = 1
    if key not in counters:
        counters[key] = start_val
    else:
        counters[key] += 1
    return counters[key]


# ============================================================================
# 后处理与清理
# ============================================================================

def _fix_chinese_bold(line: str) -> str:
    """修复中文引号粗体合并与标点移出"""
    if not line or '**' not in line:
        return line
    
    # 合并相邻的 "**xxx**" 格式
    while True:
        new_line = re.sub(r'\*\*"\*\*(.*?)\*\*"\*\*', r'**"\1"**', line)
        if new_line == line:
            break
        line = new_line
    
    # 合并相邻的粗体块
    def _merge(match):
        inner_parts = re.findall(r'\*\*([^*]+?)\*\*', match.group(0))
        return '**' + ''.join(inner_parts) + '**'
    line = re.sub(r'(?:\*\*[^*]+?\*\*){2,}', _merge, line)
    
    # 将标点符号移到粗体外面
    line = re.sub(r'\*\*(".*?")([，。！？；：,.!;:?])\*\*', r'**\1**\2', line)
    
    return line


_adjacent_bold_pattern = re.compile(r'(?:\*\*[^*]+?\*\*){2,}')
_bold_line_split_pattern = re.compile(r'^(\*\*[^*]+?\*\*)(\S.+)$')


def _merge_adjacent_bold(line: str) -> str:
    """合并相邻的粗体块"""
    def _merge(match):
        inner_parts = re.findall(r'\*\*([^*]+?)\*\*', match.group(0))
        return '**' + ''.join(inner_parts) + '**'
    return _adjacent_bold_pattern.sub(_merge, line)


def _split_leading_bold(line: str) -> list[str]:
    """分离首行完整粗体句子为独立段落"""
    # 检查是否应该分割：只有当粗体后还有很长的文本（>50字符）且以句号、叹号等结尾时才分割
    m = _bold_line_split_pattern.match(line)
    if not m:
        return [line]
    bold_block = m.group(1).rstrip()
    rest = m.group(2).lstrip()
    # 如果后续文本太短或不是完整句子，不分割
    if not rest or len(rest) < 50:
        return [line]
    # 检查后续文本是否以完整句子结束
    if not rest.rstrip()[-1] in '。！？.!?':
        return [line]
    return [bold_block, rest]


def _post_process(lines: list[str]) -> list[str]:
    """后处理清理：NBSP 替换、粗体合并、引号修复、软换行添加等"""
    processed: list[str] = []
    for line in lines:
        # 替换 NBSP
        if '\u00A0' in line:
            line = line.replace('\u00A0', ' ')
        
        # 修复粗体引号格式
        if '**' in line:
            line = fix_bold_quotes(line)
            line = _merge_adjacent_bold(line)
            
            # 将标点符号移到粗体外面
            line = re.sub(r'\*\*(".*?")([，。！？；：,.!;:?])\*\*', r'**\1**\2', line)
            # 将冒号移到粗体外面（常见于列表项标题）
            line = re.sub(r'\*\*([^*]+?)([：:])\*\*', r'**\1**\2', line)
            
            for part_line in _split_leading_bold(line):
                processed.append(part_line)
        else:
            processed.append(line)
    
    # 修复粗体/引号后紧跟破折号等符号的粘连问题，添加空格分隔
    final_processed = []
    for line in processed:
        # 处理各种引号和破折号的组合
        # U+2014 是长破折号 —, U+2013 是短破折号 –
        line = re.sub(r'(["\u201C\u201D])(\u2014{1,2}|\u2013)', r'\1 \2', line)
        line = re.sub(r'(\*\*)(\u2014{1,2}|\u2013)', r'\1 \2', line)
        final_processed.append(line)

    # 添加软换行（行尾两空格）
    final: list[str] = []
    for ln in final_processed:
        stripped = ln.strip()
        if not ln:
            final.append(ln)
            continue
        if ln.startswith('|') and ln.endswith('|'):  # 表格行
            final.append(ln)
            continue
        if stripped == '---':  # 水平线
            final.append(ln)
            continue
        if ln.startswith('!['):  # 图片引用
            final.append(ln)
            continue
        if ln.lstrip().startswith('-'):  # 列表项
            final.append(ln)
            continue
        if ln.lstrip().startswith(('#', '>')):  # 标题或引用块
            final.append(ln)
            continue
        # 只在不以标点符号结尾的行添加软换行
        if not ln.endswith('  '):
            final.append(ln + '  ')
        else:
            final.append(ln)
    return final


# ============================================================================
# 主转换函数
# ============================================================================

def extract(docx_path: str, normalize: bool = False) -> Path:
    """将 Word 文档转换为 Markdown
    
    Args:
        docx_path: Word 文档路径
        normalize: 是否规范化文件名
        
    Returns:
        生成的 Markdown 文件路径
    """
    p = Path(docx_path)
    if not p.exists():
        raise FileNotFoundError(f"文件不存在: {docx_path}")
    
    doc = Document(str(p))
    original_base = p.stem.strip()
    base_name = normalize_name(original_base) if normalize else original_base
    
    # 图片目录：顶层 Images 下，以 base_name_images 命名
    images_root = p.parent / 'Images'
    images_root.mkdir(exist_ok=True)
    images_dir = images_root / f"{base_name}_images"
    images_dir.mkdir(exist_ok=True)

    part = doc.part
    exported = {}  # 字典：rId -> 实际文件名
    md_lines = []
    numbering_counters = {}
    last_was_list = False  # 跟踪上一个段落是否为列表

    for block in _iter_block_items(doc):
        # ====== 表格处理 ======
        if isinstance(block, Table):
            table_lines = []
            for r_index, row in enumerate(block.rows):
                cells_content = []
                for cell in row.cells:
                    cell_tokens = []
                    for para in cell.paragraphs:
                        cell_tokens.extend(_runs_to_tokens(para, part, images_dir, exported))
                    cells_content.append("".join(cell_tokens))
                line = '| ' + ' | '.join(cells_content) + ' |'
                table_lines.append(line)
            if table_lines:
                # 若有多行，首行视为表头，需要分隔符行
                if len(table_lines) > 1:
                    cols = table_lines[0].count('|') - 1
                    separator = '| ' + ' | '.join(['---'] * cols) + ' |'
                    md_lines.append(table_lines[0])
                    md_lines.append(separator)
                    md_lines.extend(table_lines[1:])
                else:
                    md_lines.append(table_lines[0])
            continue

        # ====== 段落处理 ======
        paragraph = block  # type: Paragraph
        tokens = _runs_to_tokens(paragraph, part, images_dir, exported)
        if not tokens:
            # 空段落保持空行分隔
            md_lines.append("")
            continue

        style_name = paragraph.style.name if paragraph.style else ''
        
        # 处理目录样式（toc 1, toc 2, toc 3 等）
        if style_name.lower().startswith('toc '):
            try:
                level_str = style_name.split(' ')[1]
                level = int(level_str)
            except (IndexError, ValueError):
                level = 1
            
            # 获取文本并移除页码（制表符后的数字）
            text = ''.join(tokens)
            # 分离标题和页码
            if '\t' in text:
                parts = text.split('\t')
                title = parts[0].strip()
                page_num = parts[-1].strip() if len(parts) > 1 else ''
            else:
                title = text.strip()
                page_num = ''
            
            # 根据级别添加缩进
            indent = '  ' * (level - 1)
            if page_num:
                md_lines.append(f"{indent}- {title} ... {page_num}")
            else:
                md_lines.append(f"{indent}- {title}")
            continue
        
        # 标题映射
        if style_name.startswith('Heading '):
            try:
                level = int(style_name.split(' ')[1])
                level = max(1, min(level, 6))
            except ValueError:
                level = 1
            line = '#' * level + ' ' + ''.join(tokens)
            md_lines.append(line)
            continue

        # 列表段落检测与真实编号保留
        if _is_list_paragraph(paragraph):
            numFmt = paragraph._p.xpath('.//w:numPr//w:numFmt')
            if numFmt:
                val = numFmt[0].get(qn('w:val'))
                if val == 'bullet':
                    md_lines.append(f"- {''.join(tokens)}")
                    last_was_list = True
                    continue
                number = _get_list_number(paragraph, numbering_counters, part)
                if number is None:
                    number = 1
                md_lines.append(f"{number}. {''.join(tokens)}")
                last_was_list = True
                continue
            else:
                # 如果有 numPr 但没有 numFmt，默认为无序列表
                md_lines.append(f"- {''.join(tokens)}")
                last_was_list = True
                continue
        
        # 如果上一个是列表，当前不是列表，添加空行分隔
        if last_was_list:
            md_lines.append("")
            last_was_list = False
        
        # 非列表：处理段落内手动换行拆段
        if '  \n' in tokens:
            segments = []
            current = []
            for tk in tokens:
                if tk == '  \n':
                    segments.append(''.join(current))
                    current = []
                else:
                    current.append(tk)
            if current:
                segments.append(''.join(current))
            
            # 检测是否应该格式化为列表项
            should_be_list = False
            if len(segments) >= 3:
                valid_segments = [seg.strip() for seg in segments if seg.strip()]
                if len(valid_segments) >= 3:
                    starts_3 = [seg[:3] if len(seg) >= 3 else seg for seg in valid_segments]
                    starts_2 = [seg[:2] if len(seg) >= 2 else seg for seg in valid_segments]
                    ends = [seg.rstrip()[-1] if seg.rstrip() else '' for seg in valid_segments]
                    
                    from collections import Counter
                    start_counts_3 = Counter(starts_3)
                    start_counts_2 = Counter(starts_2)
                    end_counts = Counter(ends)
                    
                    most_common_start_3 = start_counts_3.most_common(1)[0] if start_counts_3 else ('', 0)
                    most_common_start_2 = start_counts_2.most_common(1)[0] if start_counts_2 else ('', 0)
                    most_common_end = end_counts.most_common(1)[0] if end_counts else ('', 0)
                    
                    # 超过一半段落有相同开头或结尾模式
                    if (most_common_start_3[1] >= len(valid_segments) / 2 or 
                        most_common_start_2[1] >= len(valid_segments) / 2 or 
                        most_common_end[1] >= len(valid_segments) / 2):
                        should_be_list = True
            
            if should_be_list:
                # 检查第一个segment是否是引导语
                first_seg = segments[0].strip() if segments else ''
                rest_segs = segments[1:] if len(segments) > 1 else []
                
                intro_prefixes = ['在这里，', '在此，', '在这里:', '在此:']
                has_intro_prefix = any(first_seg.startswith(prefix) for prefix in intro_prefixes)
                
                is_intro = False
                intro_text = ''
                first_list_item = first_seg
                
                if has_intro_prefix and len(rest_segs) >= 2:
                    # 分离引导语
                    for prefix in intro_prefixes:
                        if first_seg.startswith(prefix):
                            intro_text = prefix.rstrip('，:')
                            first_list_item = first_seg[len(prefix):].strip()
                            break
                    is_intro = True
                elif first_seg and len(rest_segs) >= 2 and len(first_seg) < 20:
                    # 第一个segment很短，视为引导语
                    is_intro = True
                    intro_text = first_seg
                    first_list_item = None
                
                if is_intro and intro_text:
                    md_lines.append(intro_text + '  ')
                    if first_list_item:
                        md_lines.append(f"- {first_list_item}")
                    for seg in rest_segs:
                        if seg.strip():
                            md_lines.append(f"- {seg}")
                    last_was_list = True
                else:
                    # 全部格式化为列表
                    for seg in segments:
                        if seg.strip():
                            md_lines.append(f"- {seg}")
                    last_was_list = True
            else:
                # 保持原样：分段处理
                for idx, seg in enumerate(segments):
                    md_lines.append(seg)
                    if idx != len(segments) - 1:
                        md_lines.append('')  # 空行分隔成新段落
        else:
            md_lines.append(''.join(tokens))

    # 回退：若未捕获图像但存在图像资源，仍需导出并在末尾引用
    if not exported:
        fallback_refs = []
        for rel_id, rel_part in part.related_parts.items():
            if hasattr(rel_part, 'content_type') and rel_part.content_type.startswith('image/'):
                filename = Path(rel_part.filename).name
                out_path = images_dir / filename
                if not out_path.exists():
                    with open(out_path, 'wb') as f:
                        f.write(rel_part.blob)
                fallback_refs.append(filename)
        for fn in sorted(fallback_refs):
            md_lines.append(f"![](Images/{images_dir.name}/{fn})")

    # 去除开头多余的空行
    while md_lines and not md_lines[0].strip():
        md_lines.pop(0)

    # 中文引号粗体合并与标点移出
    md_lines = [_fix_chinese_bold(l) for l in md_lines]

    # 后处理清理
    md_lines = _post_process(md_lines)
    
    # 修复错误的换行问题：将单独成行的冒号合并到上一行
    fixed_lines = []
    i = 0
    while i < len(md_lines):
        line = md_lines[i]
        # 检查下一行是否以冒号开头
        if i + 1 < len(md_lines) and md_lines[i + 1].strip().startswith(('：', ':')):
            # 移除当前行末尾的两个空格（软换行标记）
            if line.endswith('  '):
                line = line[:-2]
            # 合并下一行，去掉下一行开头的冒号前的空白
            next_line = md_lines[i + 1].lstrip()
            fixed_lines.append(line + next_line)
            i += 2  # 跳过下一行
        else:
            fixed_lines.append(line)
            i += 1
    
    # 写入 Markdown 文件
    md_content = "\n".join(fixed_lines) + "\n"
    md_path = p.parent / f"{base_name}.md"
    md_path.write_text(md_content, encoding='utf-8')
    return md_path


# ============================================================================
# 主程序入口
# ============================================================================

def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("使用方法: python docx_to_md.py <docx_path> [--normalize]")
        print()
        print("参数说明:")
        print("  <docx_path>   : Word 文档路径")
        print("  --normalize   : 规范化文件名（移除特殊字符，用连字符替换空格）")
        print()
        print("示例:")
        print("  python docx_to_md.py document.docx")
        print("  python docx_to_md.py document.docx --normalize")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    normalize = '--normalize' in sys.argv[2:]
    
    try:
        md_path = extract(docx_path, normalize=normalize)
        print(f"✓ 转换成功!")
        print(f"  Markdown 文件: {md_path}")
    except FileNotFoundError as e:
        print(f"✗ 错误: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"✗ 转换失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
