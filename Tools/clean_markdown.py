import re
from pathlib import Path

TARGET_EXT = ".md"

# Operations:
# 1. Replace NBSP with normal spaces
# 2. Split leading bold sentence into its own paragraph
# 3. Ensure non-empty non-table/separator lines end with two spaces for soft line break
# 4. Normalize multiple bold segments merged together (join adjacent **..** blocks)
# 5. Skip image-only lines and table separator lines for trailing spaces
# 6. Idempotent: running multiple times keeps same result

def split_leading_bold(line: str):
    """If a line starts with a bold block immediately followed by other text,
    split into two lines so the bold stands alone. Idempotent.
    Example: **Bold sentence.**Following -> [**Bold sentence.**, Following]
    """
    m = re.match(r'^(\*\*[^*]+?\*\*)(\S.+)$', line)
    if not m:
        return [line]
    bold_block = m.group(1).rstrip()
    rest = m.group(2).lstrip()
    # If rest itself is bold only we keep original
    return [bold_block, rest] if rest else [line]

_adjacent_bold_pattern = re.compile(r'(?:\*\*[^*]+?\*\*){2,}')

def merge_adjacent_bold(line: str) -> str:
    def _merge(match):
        inner_parts = re.findall(r'\*\*([^*]+?)\*\*', match.group(0))
        return '**' + ''.join(inner_parts) + '**'
    return _adjacent_bold_pattern.sub(_merge, line)

# Chinese quote punctuation fix inside bold
_chinese_punct_pattern = re.compile(r'\*\*(“.*?”)([，。！？；：,.!;:?])\*\*')

def fix_chinese_punct(line: str) -> str:
    return _chinese_punct_pattern.sub(r'**\1**\2', line)


def process_lines(lines: list[str]) -> list[str]:
    out = []
    for line in lines:
        if '\u00A0' in line:
            line = line.replace('\u00A0', ' ')
        line = merge_adjacent_bold(line)
        line = fix_chinese_punct(line)
        split_parts = split_leading_bold(line) if '**' in line else [line]
        out.extend(split_parts)

    final = []
    for ln in out:
        # Skip modification for table separator or empty line
        if ln.startswith('|') and ln.endswith('|'):
            final.append(ln)
            continue
        if ln.strip() == '---':  # horizontal rule keep as-is
            final.append(ln)
            continue
        if ln.startswith('!['):  # image reference keep
            final.append(ln)
            continue
        if ln and not ln.endswith('  '):
            ln += '  '
        final.append(ln)
    return final


def clean_file(path: Path):
    original = path.read_text(encoding='utf-8').splitlines()
    cleaned = process_lines(original)
    content = '\n'.join(cleaned) + '\n'
    path.write_text(content, encoding='utf-8')
    return True


def main():
    root = Path(__file__).resolve().parent.parent
    md_files = [p for p in root.iterdir() if p.is_file() and p.suffix == TARGET_EXT]
    if not md_files:
        print('No markdown files found at root to clean.')
        return
    for md in md_files:
        clean_file(md)
        print('Cleaned:', md.name)

if __name__ == '__main__':
    main()
