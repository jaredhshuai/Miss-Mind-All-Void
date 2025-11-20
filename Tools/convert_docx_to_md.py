import sys
import re
from pathlib import Path
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn

"""Convert a .docx file to markdown preserving paragraph order and exporting images.
Images are referenced with original filenames and empty alt text only.

Optional normalization (--normalize) will sanitize the base name:
    1. Strip leading/trailing whitespace
    2. Remove chars: 《 》 ： : “ ” "
    3. Collapse internal whitespace to single hyphen
    4. Remove trailing hyphens

Extra newline handling:
    - Manual <w:br/> breaks become explicit '  \n' tokens
    - Embedded control chars (\r, \n, \x0b, \x0c, \u2028, \u2029) are split
      and converted to the same break tokens so paragraphs are properly
      separated instead of silently merged.
"""

def normalize_name(name: str) -> str:
        # Remove specific punctuation that can cause rendering or path issues
        name = re.sub(r'[《》：“”"\:]', '', name)
        # Replace any whitespace sequence with single hyphen
        name = re.sub(r'\s+', '-', name.strip())
        # Collapse multiple hyphens
        name = re.sub(r'-{2,}', '-', name)
        return name.strip('-')

def _extract_run_style(run):
    """提取单个 run 的文本与样式，不直接返回已包裹的 markdown，方便后续合并。"""
    return {
        'text': run.text or '',
        'bold': bool(run.bold),
        'italic': bool(run.italic),
        'is_image': False,
        'is_line_break': False,
    }

_EXTRA_BREAK_CHARS_PATTERN = re.compile(r'[\r\n\u000b\u000c\u2028\u2029]')

def _split_text_with_breaks(text: str) -> list:
    """将包含额外换行控制符的文本拆分为纯文本和显式换行 token。
    使用与手动 <w:br/> 相同的 '  \n' 作为换行标记，后续逻辑即可复用。
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
    # 若文本以换行结束，最后一个换行已经作为 token 添加，无需额外空 token
    return parts

def _format_merged(text: str, bold: bool, italic: bool) -> str:
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
    """Yield paragraphs and tables in document order."""
    for child in doc.element.body:
        if child.tag.endswith('p'):
            yield Paragraph(child, doc)
        elif child.tag.endswith('tbl'):
            yield Table(child, doc)

def _runs_to_tokens(paragraph, part, images_dir: Path, exported: set) -> list:
    # 采集结构化 run
    raw_runs = []
    for run in getattr(paragraph, 'runs', []):
        data = _extract_run_style(run)
        inline = run.element
        # 处理内部手动换行 <w:br/>
        if inline.xpath('.//w:br'):
            raw_runs.append({'text': '  \n', 'bold': False, 'italic': False, 'is_image': False, 'is_line_break': True})
        # 图像
        blips = inline.xpath('.//a:blip')
        for blip in blips:
            rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if not rId:
                continue
            image_part = part.related_parts.get(rId)
            if image_part is None:
                continue
            filename = Path(image_part.filename).name
            out_path = images_dir / filename
            if filename not in exported:
                with open(out_path, 'wb') as f:
                    f.write(image_part.blob)
                exported.add(filename)
            raw_runs.append({'text': f"![](Images/{images_dir.name}/{filename})", 'bold': False, 'italic': False, 'is_image': True, 'is_line_break': False})
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

    # 合并连续样式一致（且非图片/换行） 的 run，减少重复 ** 包裹
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
    for m in merged:
        if m['is_image'] or m['is_line_break']:
            tokens.append(m['text'])
        else:
            formatted = _format_merged(m['text'], m['bold'], m['italic'])
            if formatted:
                tokens.append(formatted)
    return tokens

def _is_list_paragraph(paragraph) -> bool:
    return bool(paragraph._p.xpath('.//w:numPr'))

def _list_marker(paragraph) -> str:
    numFmt = paragraph._p.xpath('.//w:numPr//w:numFmt')
    if numFmt:
        val = numFmt[0].get(qn('w:val'))
        if val == 'bullet':
            return '- '
        return '1. '
    return '- '

def _get_list_number(paragraph, counters: dict, part) -> int | None:
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
        start_e = numbering.xpath(f'.//w:abstractNum[w:abstractNumId[@w:val="{abstract_id}"]]//w:lvl[@w:ilvl="{ilvl}"]//w:start')
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

def extract(docx_path: str, normalize: bool = False) -> Path:
    p = Path(docx_path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {docx_path}")
    doc = Document(str(p))
    original_base = p.stem.strip()
    base_name = normalize_name(original_base) if normalize else original_base
    # 图片目录改为顶层 Images 下，以 base_name_images 命名
    images_root = p.parent / 'Images'
    images_root.mkdir(exist_ok=True)
    images_dir = images_root / f"{base_name}_images"
    images_dir.mkdir(exist_ok=True)

    part = doc.part
    exported = set()
    md_lines = []
    numbering_counters = {}

    for block in _iter_block_items(doc):
        # 表格处理
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
                # 若有多行，首行视为表头，需要分隔符行（Markdown 语法必要，不属于额外内容）
                if len(table_lines) > 1:
                    cols = table_lines[0].count('|') - 1
                    separator = '| ' + ' | '.join(['---'] * cols) + ' |'
                    md_lines.append(table_lines[0])
                    md_lines.append(separator)
                    md_lines.extend(table_lines[1:])
                else:
                    md_lines.append(table_lines[0])
            continue

        # 段落处理
        paragraph = block  # type: Paragraph
        tokens = _runs_to_tokens(paragraph, part, images_dir, exported)
        if not tokens:
            # 空段落保持空行分隔
            md_lines.append("")
            continue

        style_name = paragraph.style.name if paragraph.style else ''
        # 标题映射
        if style_name.startswith('Heading '):
            try:
                level = int(style_name.split(' ')[1])
                level = max(1, min(level, 6))
            except ValueError:
                level = 1
            line = '#' * level + ' ' + ''.join(tokens)
            md_lines.append(line)
            # 结束任何进行中的有序列表
            list_sequence = 0
            in_number_list = False
            continue

        # 列表段落检测与真实编号保留
        if _is_list_paragraph(paragraph):
            numFmt = paragraph._p.xpath('.//w:numPr//w:numFmt')
            if numFmt:
                val = numFmt[0].get(qn('w:val'))
                if val == 'bullet':
                    md_lines.append(f"- {''.join(tokens)}")
                    continue
                number = _get_list_number(paragraph, numbering_counters, part)
                if number is None:
                    number = 1
                md_lines.append(f"{number}. {''.join(tokens)}")
                continue
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
            for idx, seg in enumerate(segments):
                md_lines.append(seg)
                if idx != len(segments) - 1:
                    md_lines.append('')  # 空行分隔成新段落
        else:
            md_lines.append(''.join(tokens))

    # 回退：若未捕获图像但存在图像资源，仍需导出并在末尾引用，避免遗漏
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

    # 去除开头多余的空行，避免首行即空行导致渲染不一致
    while md_lines and not md_lines[0].strip():
        md_lines.pop(0)

    # 中文引号粗体合并与标点移出
    def _fix_chinese_bold(line: str) -> str:
        if not line or '**' not in line:
            return line
        while True:
            new_line = re.sub(r'\*\*“\*\*(.*?)\*\*”\*\*', r'**“\1”**', line)
            if new_line == line:
                break
            line = new_line
        def _merge(match):
            inner_parts = re.findall(r'\*\*([^*]+?)\*\*', match.group(0))
            return '**' + ''.join(inner_parts) + '**'
        line = re.sub(r'(?:\*\*[^*]+?\*\*){2,}', _merge, line)
        line = re.sub(r'\*\*(“.*?”)([，。！？；：,.!;:?])\*\*', r'**\1**\2', line)
        return line
    md_lines = [_fix_chinese_bold(l) for l in md_lines]
    md_content = "\n".join(md_lines) + "\n"
    md_path = p.parent / f"{base_name}.md"
    md_path.write_text(md_content, encoding='utf-8')
    return md_path

def main():
    if len(sys.argv) < 2:
        print("Usage: python convert_docx_to_md.py <docx_path> [--normalize]")
        sys.exit(1)
    docx_path = sys.argv[1]
    normalize = '--normalize' in sys.argv[2:]
    md_path = extract(docx_path, normalize=normalize)
    print("Markdown created:", md_path)

if __name__ == "__main__":
    main()
