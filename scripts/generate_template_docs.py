# -*- coding: utf-8 -*-
"""
학습일지 표준 작성 템플릿 DOCX 생성기
references/학습일지_표준_작성_템플릿.md 내용을 바탕으로 정형화된 Word(.docx) 문서를 생성합니다.
"""
import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def create_styled_docx(md_path, docx_path):
    doc = docx.Document()
    
    # 페이지 여백 설정 (기본 1인치 -> 상하좌우 20mm)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # 기본 스타일 폰트 설정
    style = doc.styles['Normal']
    font = style.font
    font.name = '맑은 고딕'
    font.size = Pt(10)
    font.color.rgb = RGBColor(0x1e, 0x29, 0x3b)

    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    in_table = False
    table_lines = []

    def flush_table(tbl_lines):
        if not tbl_lines:
            return
        parsed_rows = []
        for l in tbl_lines:
            parts = [p.strip() for p in l.strip().split('|')[1:-1]]
            if all(set(p).issubset({'-', ':', ' '}) for p in parts):
                continue  # 구분선
            parsed_rows.append(parts)

        if not parsed_rows:
            return

        cols = max(len(r) for r in parsed_rows)
        tbl = doc.add_table(rows=len(parsed_rows), cols=cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        for r_idx, row in enumerate(parsed_rows):
            for c_idx in range(cols):
                cell = tbl.cell(r_idx, c_idx)
                val = row[c_idx] if c_idx < len(row) else ""
                cell.text = val
                set_cell_margins(cell, top=120, bottom=120, left=150, right=150)
                
                # 폰트 스타일링
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.space_before = Pt(2)
                    for run in p.runs:
                        run.font.name = '맑은 고딕'
                        run.font.size = Pt(9)
                        if r_idx == 0:
                            run.font.bold = True
                            run.font.color.rgb = RGBColor(0x0f, 0x17, 0x2a)

                # 헤더 음영
                if r_idx == 0:
                    set_cell_background(cell, "F1F5F9")
                elif r_idx % 2 == 1:
                    set_cell_background(cell, "FFFFFF")
                else:
                    set_cell_background(cell, "F8FAFC")

        # 테이블 아래 여백
        doc.add_paragraph()

    i = 0
    while i < len(lines):
        line = lines[i].rstrip('\r\n')

        # 테이블 감지
        if line.strip().startswith('|') and '|' in line[1:]:
            table_lines.append(line)
            i += 1
            continue
        elif table_lines:
            flush_table(table_lines)
            table_lines = []

        # 제목 (#, ##, ###)
        if line.startswith('# '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(line[2:].strip())
            run.font.name = '맑은 고딕'
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0x1e, 0x3a, 0x8a)
        elif line.startswith('## '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(line[3:].strip())
            run.font.name = '맑은 고딕'
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0x1e, 0x40, 0xaf)
        elif line.startswith('### '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(line[4:].strip())
            run.font.name = '맑은 고딕'
            run.font.size = Pt(11)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0x0f, 0x17, 0x2a)
        elif line.startswith('> '):
            # 안내 박스
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.left_indent = Inches(0.2)
            run = p.add_run(line[2:].strip())
            run.font.name = '맑은 고딕'
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(0x33, 0x41, 0x55)
            if '⚠️' in line or '핵심' in line:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x9a, 0x34, 0x12)
        elif line.strip() == '---':
            # 구분선
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run('―' * 45)
            run.font.color.rgb = RGBColor(0xcb, 0xd5, 0xe1)
        elif line.strip():
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(line)
            run.font.name = '맑은 고딕'
            run.font.size = Pt(10)

        i += 1

    if table_lines:
        flush_table(table_lines)

    doc.save(docx_path)
    print(f"생성 완료: {docx_path}")

if __name__ == '__main__':
    md = 'c:/workspace/학습일지 평가/references/학습일지_표준_작성_템플릿.md'
    out1 = 'c:/workspace/학습일지 평가/references/학습일지_표준_작성_템플릿.docx'
    out2 = 'c:/workspace/학습일지 평가/references/경혈학_학습일지_표준_작성_템플릿.docx'
    create_styled_docx(md, out1)
    create_styled_docx(md, out2)
