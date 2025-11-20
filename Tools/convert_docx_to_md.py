import sys
import re
from pathlib import Path
from docx import Document

"""Convert a .docx file to markdown preserving paragraph order and exporting images.
Images are referenced with original filenames and empty alt text only.

Optional normalization (--normalize) will sanitize the base name:
    1. Strip leading/trailing whitespace
    2. Remove chars: 《 》 ： : “ ” "
    3. Collapse internal whitespace to single hyphen
    4. Remove trailing hyphens
"""

def normalize_name(name: str) -> str:
        # Remove specific punctuation that can cause rendering or path issues
        name = re.sub(r'[《》：“”"\:]', '', name)
        # Replace any whitespace sequence with single hyphen
        name = re.sub(r'\s+', '-', name.strip())
        # Collapse multiple hyphens
        name = re.sub(r'-{2,}', '-', name)
        return name.strip('-')

def _format_run_text(run) -> str:
    """Return markdown formatted text for a single run preserving bold/italic.
    不额外添加内容：仅根据原有样式加包裹符号。组合粗体+斜体用三星。"""
    text = run.text
    if not text:
        return ''
    bold = bool(run.bold)
    italic = bool(run.italic)
    # Underline 不转成 markdown，下划线容易引入歧义，保持原文字即可。
    if bold and italic:
        return f"***{text}***"
    if bold:
        return f"**{text}**"
    if italic:
        return f"*{text}*"
    return text

def extract(docx_path: str, normalize: bool = False) -> Path:
    p = Path(docx_path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {docx_path}")
    doc = Document(str(p))
    original_base = p.stem.strip()
    base_name = normalize_name(original_base) if normalize else original_base
    images_dir = p.parent / f"{base_name}_images"
    images_dir.mkdir(exist_ok=True)

    part = doc.part
    exported = set()
    md_lines = []

    # 遍历段落，按顺序构建段落内容（文本与图像占位）
    for paragraph in doc.paragraphs:
        segment_tokens = []
        for run in paragraph.runs:
            inline = run.element
            # 处理图像（可能在同一段落的任意位置）
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
                segment_tokens.append(f"![]({images_dir.name}/{filename})")
            # 处理文本格式
            formatted = _format_run_text(run)
            if formatted:
                segment_tokens.append(formatted)
        # 仅当段落有内容（图像或文本）才加入
        if segment_tokens:
            # 将同一段落的内容合并为单行，保持原序，不额外添加空格
            md_lines.append("".join(segment_tokens))

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
            md_lines.append(f"![]({images_dir.name}/{fn})")

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
