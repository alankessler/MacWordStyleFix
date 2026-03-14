#!/usr/bin/env python3
"""
docx-fix-styles — fix Word style picker rendering issues in .docx files.

Fixes two problems caused by Pages exports (and some older Word versions):

1. w14:textOutline blocks — verbose "no outline" markers on every style
   definition that bloat the XML and can confuse Word's renderer.

2. Normal style excessive line spacing — Pages sets atLeast 455 twips
   (32pt minimum) for 9pt text. Word's style picker preview rows are
   too short for this, causing all style names to be clipped/unreadable.

Usage:
    docx-fix-styles <file.docx> [-o output.docx]
"""

import sys
import os
import re
import zipfile
import tempfile
import argparse


_TEXTOUTLINE_RE = re.compile(
    r'<w14:textOutline[^>]*>.*?</w14:textOutline>', re.DOTALL)

# Match <w:spacing> with w:lineRule="atLeast" and a w:line value,
# regardless of attribute order or other attributes present.
_SPACING_ATLEAST_RE = re.compile(
    r'<w:spacing\b(?=[^>]*w:lineRule="atLeast")(?=[^>]*w:line="(\d+)")[^>]*/>')


def fix_styles(xml_text):
    """Fix style definitions that break Word's style picker."""
    fixes = []

    # Remove w14:textOutline blocks
    xml_text, count = _TEXTOUTLINE_RE.subn('', xml_text)
    if count:
        fixes.append(f"Removed {count} w14:textOutline block(s)")

    # Fix Normal style's excessive line spacing
    m = re.search(
        r'<w:style[^>]*w:styleId="Normal"[^>]*>.*?</w:style>',
        xml_text, re.DOTALL)
    if m:
        style_block = m.group(0)
        spacing_m = _SPACING_ATLEAST_RE.search(style_block)
        if spacing_m and int(spacing_m.group(1)) > 300:
            old_val = spacing_m.group(1)
            fixes.append(
                f"Fixed Normal line spacing: {old_val} twips atLeast \u2192 240 auto")
            new_spacing = spacing_m.group(0)
            new_spacing = re.sub(r'w:line="\d+"', 'w:line="240"', new_spacing)
            new_spacing = new_spacing.replace(
                'w:lineRule="atLeast"', 'w:lineRule="auto"')
            new_block = style_block.replace(
                spacing_m.group(0), new_spacing)
            xml_text = xml_text[:m.start()] + new_block + xml_text[m.end():]

    return xml_text, fixes


def main():
    parser = argparse.ArgumentParser(
        description="Fix Word style picker rendering issues in .docx files.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument("file", help="Path to .docx file")
    parser.add_argument("-o", "--output",
                        help="Output file (default: overwrite input)")
    args = parser.parse_args()

    if not os.path.exists(args.file):
        print(f"Error: {args.file} not found", file=sys.stderr)
        sys.exit(1)

    try:
        zin = zipfile.ZipFile(args.file, 'r')
    except zipfile.BadZipFile:
        print(f"Error: {args.file} is not a valid .docx (zip) file",
              file=sys.stderr)
        sys.exit(1)

    dst = args.output or args.file

    # Read styles.xml from the archive
    with zin:
        if "word/styles.xml" not in zin.namelist():
            print("No word/styles.xml found — nothing to fix.")
            return

        original = zin.read("word/styles.xml").decode('utf-8')
        xml_text, fixes = fix_styles(original)

        if xml_text == original:
            print("Style definitions look clean — nothing to fix.")
            return

        # Repack: copy all entries from original zip, replacing only styles.xml
        tmp_fd, tmp_path = tempfile.mkstemp(
            suffix='.docx', dir=os.path.dirname(os.path.abspath(dst)))
        try:
            os.close(tmp_fd)
            with zipfile.ZipFile(tmp_path, 'w') as zout:
                for item in zin.infolist():
                    if item.filename == "word/styles.xml":
                        zout.writestr(item, xml_text.encode('utf-8'))
                    else:
                        zout.writestr(item, zin.read(item.filename))
            os.replace(tmp_path, dst)
        except BaseException:
            os.unlink(tmp_path)
            raise

    for fix in fixes:
        print(f"  {fix}")
    print(f"Output: {dst}")


if __name__ == "__main__":
    main()
