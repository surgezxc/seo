#!/usr/bin/env python3
"""Собирает docx с ответами разработчику по заданию от 22 ноября.

Запуск:
  /Users/vladislavkhodarev/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/bin/python3 \
      tools/make-answers-docx.py

Нужен, потому что Владимир не имеет доступа к репозиторию: ответы должны
уехать одним файлом. Источник — docs/Ответы-на-вопросы_2026-11-24.md.
"""
import os
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm, Pt

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, "docs", "Ответы-на-вопросы_2026-11-24.md")
OUT = os.path.join(BASE, "docs", "Ответы-на-вопросы_2026-11-24.docx")

INLINE = re.compile(r"(\*\*.+?\*\*|`[^`]+`)")


def read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def strip_q(s):
    return re.sub(r"^>\s?", "", s).rstrip()


def balance_bold(md):
    """Склеивает строки, где ** открылся на одной строке, а закрылся на другой."""
    out, buf = [], ""
    in_code = False
    for line in md.split("\n"):
        if line.strip().startswith("```"):
            if buf:
                out.append(buf)
                buf = ""
            out.append(line)
            in_code = not in_code
            continue
        if in_code:
            out.append(line)
            continue
        buf = (buf + " " + line.strip()) if buf else line
        if buf.count("**") % 2 == 0 and buf.count("`") % 2 == 0:
            out.append(buf)
            buf = ""
    if buf:
        out.append(buf)
    return "\n".join(out)


def add_runs(p, text):
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", text)
    for part in INLINE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            r = p.add_run(part[2:-2].replace("`", ""))
            r.bold = True
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            r = p.add_run(part[1:-1])
            r.font.name = "Consolas"
            r.font.size = Pt(9.5)
        else:
            p.add_run(part.replace("`", ""))


def is_table_start(lines, i):
    if i + 1 >= len(lines):
        return False
    a, b = strip_q(lines[i]).strip(), strip_q(lines[i + 1]).strip()
    return a.startswith("|") and re.match(r"^\|[\s:|-]+\|$", b) is not None


doc = Document()
st = doc.styles["Normal"]
st.font.name = "Times New Roman"
st.font.size = Pt(11)
st.paragraph_format.space_after = Pt(4)
st.paragraph_format.line_spacing = 1.15


def flush_block(block):
    first = block[0]
    raw = first.strip()

    if raw in ("---", "***", "___"):
        return

    # Цитата: строки блока начинаются с ">".
    if all(l.strip().startswith(">") for l in block):
        text = " ".join(strip_q(l).strip() for l in block)
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.space_after = Pt(2)
        add_runs(p, text)
        for r in p.runs:
            r.italic = True
        return

    m_head = re.match(r"^(#{1,4}) (.*)$", raw)
    if m_head:
        add_runs(doc.add_heading(level=len(m_head.group(1))), m_head.group(2))
        return

    # Списки: склеиваем переносы строк в один пункт.
    m_bullet = re.match(r"^\s*[-*] (.*)$", first)
    if m_bullet:
        text = " ".join([m_bullet.group(1)] + [l.strip() for l in block[1:]])
        add_runs(doc.add_paragraph(style="List Bullet"), text)
        return

    m_num = re.match(r"^\s*(\d+)\. (.*)$", first)
    if m_num:
        # Нумерацию пишем текстом: стиль List Number в Word продолжает
        # нумерацию через весь документ, и второй список начинается с 9.
        text = f"{m_num.group(1)}. " + " ".join(
            [m_num.group(2)] + [l.strip() for l in block[1:]]
        )
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.first_line_indent = Cm(-0.6)
        add_runs(p, text)
        return

    text = " ".join(l.strip() for l in block)
    add_runs(doc.add_paragraph(), text)


def render(md):
    lines = balance_bold(md).split("\n")
    i = 0
    while i < len(lines):
        s = lines[i].rstrip()
        if not s.strip():
            i += 1
            continue

        if s.startswith("```"):
            i += 1
            buf = []
            while i < len(lines) and not lines[i].startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1
            for b in buf:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(b if b else " ")
                r.font.name = "Consolas"
                r.font.size = Pt(8.5)
            doc.add_paragraph()
            continue

        if is_table_start(lines, i):
            header = [c.strip() for c in s.strip("|").split("|")]
            i += 2
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            t = doc.add_table(rows=1, cols=len(header))
            t.style = "Table Grid"
            t.alignment = WD_TABLE_ALIGNMENT.CENTER
            for j, h in enumerate(header):
                cell = t.rows[0].cells[j]
                cell.text = ""
                add_runs(cell.paragraphs[0], h)
                for r in cell.paragraphs[0].runs:
                    r.bold = True
                    r.font.size = Pt(9)
            for row in rows:
                cells = t.add_row().cells
                for j, c in enumerate(row[: len(header)]):
                    cells[j].text = ""
                    add_runs(cells[j].paragraphs[0], c)
                    for r in cells[j].paragraphs[0].runs:
                        r.font.size = Pt(9)
            doc.add_paragraph()
            continue

        is_list = re.match(r"^\s*(?:[-*]|\d+\.)\s", s) is not None
        block = [s]
        i += 1
        while i < len(lines) and lines[i].strip():
            nxt = lines[i]
            if nxt.startswith("```") or is_table_start(lines, i) or re.match(r"^#{1,4} ", nxt):
                break
            if is_list:
                if re.match(r"^\s*(?:[-*]|\d+\.)\s", nxt):
                    break
                if not nxt.startswith((" ", "\t")):
                    break
            elif re.match(r"^\s*(?:[-*]|\d+\.)\s", nxt):
                break
            block.append(nxt)
            i += 1
        flush_block(block)


render(read(SRC))
doc.save(OUT)
print("saved", OUT, os.path.getsize(OUT) // 1024, "KB")
