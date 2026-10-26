#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_source_docx.py (V1.0 最终版)

Generates the required source code Word (.docx) document for Chinese software copyright applications.
Strictly implements the standards in 《软著申请源代码指南.md》:
1. Page setup: A4, Top 2.5cm, Bottom 2.5cm, Left 2.0cm, Right 2.0cm, Header 1.5cm, Footer 1.5cm.
2. Header: 1x2 borderless table spanning 17.0cm (9638 dxa), left-aligned title, right-aligned PAGE/NUMPAGES dynamic fields, 0.75pt bottom line.
3. Typography: Song / SimSun 8.5 pt (小五号), exactly 12.0 pt line spacing, space before/after 0 pt.
4. Natural code flow: No manual page breaks.
5. Slicing (节选) standard: If total lines > 3500 lines, automatically slices to front 1700 lines + "// [... 中间部分已省略 ...]" + rear 1700 lines (exactly 60 pages).
   If total lines <= 3500 lines, preserves all lines (full submission for < 60 pages per copyright rule 10).
"""

import argparse
import os
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn
except ImportError:
    sys.stderr.write("Error: python-docx is required. Install via `pip install -r requirements-optional.txt`\n")
    sys.exit(1)


def build_source_docx(
    input_text_path: Path,
    output_docx_path: Path,
    software_name: str,
    software_version: str = "V1.0",
    max_lines_slice: int = 3500,
) -> Path:
    """Builds a standardized source code docx file from a cleaned source text file."""
    if not input_text_path.exists():
        raise FileNotFoundError(f"Input text file not found: {input_text_path}")

    with open(input_text_path, "r", encoding="utf-8", errors="ignore") as f:
        raw_lines = f.readlines()

    total_lines = len(raw_lines)
    if total_lines > max_lines_slice:
        # Slice into front 1700 + separator + rear 1700 lines -> exactly 60 pages
        half = 1700
        code_lines = raw_lines[:half] + ["// [... 中间部分已省略 ...]\n"] + raw_lines[-half:]
    else:
        code_lines = raw_lines

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    sec.header_distance = Cm(1.5)
    sec.footer_distance = Cm(1.5)
    sec.different_first_page_header_footer = False

    # Header: 1x2 borderless table, 17.0cm = 9638 dxa (left 6236 dxa, right 3402 dxa)
    header = sec.header
    for p in list(header.paragraphs):
        p._element.getparent().remove(p._element)

    title_text = f"{software_name.strip()} {software_version.strip()} 源代码".strip()
    header_tbl_xml = parse_xml(f'''
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:tblPr>
                <w:tblW w:type="dxa" w:w="9638"/>
                <w:tblCellMar>
                    <w:top w:w="0" w:type="dxa"/>
                    <w:bottom w:w="40" w:type="dxa"/>
                    <w:left w:w="0" w:type="dxa"/>
                    <w:right w:w="0" w:type="dxa"/>
                </w:tblCellMar>
                <w:tblBorders>
                    <w:top w:val="none"/>
                    <w:left w:val="none"/>
                    <w:bottom w:val="single" w:sz="6" w:space="1" w:color="999999"/>
                    <w:right w:val="none"/>
                    <w:insideH w:val="none"/>
                    <w:insideV w:val="none"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="6236"/>
                <w:gridCol w:w="3402"/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:tcPr>
                        <w:tcW w:type="dxa" w:w="6236"/>
                        <w:tcBorders>
                            <w:bottom w:val="single" w:sz="6" w:space="1" w:color="999999"/>
                        </w:tcBorders>
                    </w:tcPr>
                    <w:p>
                        <w:pPr>
                            <w:spacing w:after="40" w:line="240" w:lineRule="auto"/>
                        </w:pPr>
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/>
                                <w:sz w:val="18"/>
                                <w:szCs w:val="18"/>
                            </w:rPr>
                            <w:t>{title_text}</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr>
                        <w:tcW w:type="dxa" w:w="3402"/>
                        <w:tcBorders>
                            <w:bottom w:val="single" w:sz="6" w:space="1" w:color="999999"/>
                        </w:tcBorders>
                    </w:tcPr>
                    <w:p>
                        <w:pPr>
                            <w:spacing w:after="40" w:line="240" w:lineRule="auto"/>
                            <w:jc w:val="right"/>
                        </w:pPr>
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/>
                                <w:sz w:val="18"/>
                                <w:szCs w:val="18"/>
                            </w:rPr>
                            <w:t xml:space="preserve">第 </w:t>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/>
                                <w:sz w:val="18"/>
                                <w:szCs w:val="18"/>
                            </w:rPr>
                            <w:fldChar w:fldCharType="begin"/>
                            <w:instrText xml:space="preserve">PAGE</w:instrText>
                            <w:fldChar w:fldCharType="separate"/>
                            <w:fldChar w:fldCharType="end"/>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/>
                                <w:sz w:val="18"/>
                                <w:szCs w:val="18"/>
                            </w:rPr>
                            <w:t xml:space="preserve"> 页 共 </w:t>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/>
                                <w:sz w:val="18"/>
                                <w:szCs w:val="18"/>
                            </w:rPr>
                            <w:fldChar w:fldCharType="begin"/>
                            <w:instrText xml:space="preserve">NUMPAGES</w:instrText>
                            <w:fldChar w:fldCharType="separate"/>
                            <w:fldChar w:fldCharType="end"/>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/>
                                <w:sz w:val="18"/>
                                <w:szCs w:val="18"/>
                            </w:rPr>
                            <w:t xml:space="preserve"> 页</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
            </w:tr>
        </w:tbl>
    ''')
    header._element.append(header_tbl_xml)

    # Append code lines
    for line in code_lines:
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p.paragraph_format.line_spacing = Pt(12)

        text = line.rstrip("\r\n")
        r = p.add_run(text)
        r.font.name = "SimSun"
        r.font.size = Pt(8.5)
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")

    output_docx_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx_path))
    return output_docx_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build editable DOCX source code document for Chinese copyright registration (V1.0 Final Edition)."
    )
    parser.add_argument("--input", required=True, help="Path to cleaned input source text file.")
    parser.add_argument("--output", required=True, help="Path to output .docx file.")
    parser.add_argument("--name", required=True, help="Software name / full name.")
    parser.add_argument("--version", default="V1.0", help="Software version (default: V1.0).")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)
    if output_path.exists():
        sys.stderr.write(f"Error: Output file already exists: {output_path.name}\n")
        sys.exit(1)

    build_source_docx(input_path, output_path, args.name, args.version)
    print(f"Standardized source code DOCX generated: {output_path}")


if __name__ == "__main__":
    main()
