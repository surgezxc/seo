#!/usr/bin/env python3
"""Собирает один docx для Владимира: задание + тексты шести страниц.

Запуск:
  /Users/vladislavkhodarev/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/bin/python3 \
      tools/make-vladimir-docx.py

Зачем скрипт, а не ручная сборка: Владимир не имеет доступа к репозиторию,
поэтому тексты должны лежать внутри документа. Пересобирается одной командой
после любых правок в docs/.
"""
import re, os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "docs", "Задание-по-сайту-2026-11-18.docx")

FILES = [
    ("Приложение А. Страница «Проверка объекта перед покупкой»", "docs/text-proverka-obekta_2026-11-18.md"),
    ("Приложение Б. Страница «Юристы»", "docs/text-yuristy_2026-11-18.md"),
    ("Приложение В. Страница «Сделки с недвижимостью»", "docs/text-sdelki_2026-11-18.md"),
    ("Приложение Г. Страница «Анализ коммерческой недвижимости»", "docs/text-kommercia_2026-11-18.md#1"),
    ("Приложение Д. Страница «Продажа коммерческой недвижимости»", "docs/text-kommercia_2026-11-18.md#2"),
    ("Приложение Е. Страница «Запуск агентства недвижимости»", "docs/text-zapusk-agentstva_2026-11-18.md"),
]


def read(rel):
    with open(os.path.join(BASE, rel), encoding="utf-8") as f:
        return f.read()


def balance_bold(md):
    """Склеивает строки, где ** или ` открылся на одной строке, а закрылся на
    другой. Блоки кода (```) не трогает."""
    out, buf = [], ""
    in_code = False
    for line in md.split("\n"):
        if line.strip().startswith("```"):
            if buf:
                out.append(buf); buf = ""
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


def strip_q(s):
    return re.sub(r"^>\s?", "", s).rstrip()


INLINE = re.compile(r"(\*\*.+?\*\*|`[^`]+`)")


def add_runs(p, text):
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", text)
    for part in INLINE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            r = p.add_run(part[2:-2].replace("`", "")); r.bold = True
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            r = p.add_run(part[1:-1]); r.font.name = "Consolas"; r.font.size = Pt(9.5)
        else:
            p.add_run(part.replace("`", ""))


def is_table_start(lines, i):
    if i + 1 >= len(lines):
        return False
    a, b = strip_q(lines[i]).strip(), strip_q(lines[i + 1]).strip()
    return a.startswith("|") and re.match(r"^\|[\s:|-]+\|$", b) is not None


def render(md):
    lines = balance_bold(md).split("\n")
    i = 0
    while i < len(lines):
        s = lines[i].rstrip()
        raw = strip_q(s).strip()

        if is_table_start(lines, i):
            header = [c.strip() for c in raw.strip("|").split("|")]
            i += 2
            rows = []
            while i < len(lines) and strip_q(lines[i]).strip().startswith("|"):
                rows.append([c.strip() for c in strip_q(lines[i]).strip().strip("|").split("|")])
                i += 1
            t = doc.add_table(rows=1, cols=len(header))
            t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
            for j, h in enumerate(header):
                cell = t.rows[0].cells[j]; cell.text = ""
                add_runs(cell.paragraphs[0], h)
                for r in cell.paragraphs[0].runs:
                    r.bold = True; r.font.size = Pt(9)
            for row in rows:
                cells = t.add_row().cells
                for j, c in enumerate(row[:len(header)]):
                    cells[j].text = ""
                    add_runs(cells[j].paragraphs[0], c)
                    for r in cells[j].paragraphs[0].runs:
                        r.font.size = Pt(9)
            doc.add_paragraph()
            continue

        if s.startswith("```"):
            i += 1
            buf = []
            while i < len(lines) and not lines[i].startswith("```"):
                buf.append(lines[i]); i += 1
            i += 1
            for b in buf:
                p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0)
                r = p.add_run(b if b else " "); r.font.name = "Consolas"; r.font.size = Pt(8.5)
            doc.add_paragraph()
            continue

        if s.startswith(">"):
            m_h = re.match(r"^\*\*\[(H[12])\]\s*(.+?)\*\*$", raw)
            if m_h:
                add_runs(doc.add_heading(level=2 if m_h.group(1) == "H1" else 3),
                         m_h.group(2))
                i += 1
                continue
            m_li = re.match(r"^\s*[-*]\s+(.*)$", raw)
            if m_li:
                add_runs(doc.add_paragraph(style="List Bullet"), m_li.group(1))
                i += 1
                continue
            m_ol = re.match(r"^\s*\d+\.\s+(.*)$", raw)
            if m_ol:
                add_runs(doc.add_paragraph(style="List Number"), m_ol.group(1))
                i += 1
                continue
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.6)
            p.paragraph_format.space_after = Pt(2)
            add_runs(p, raw)
            for r in p.runs:
                r.italic = True
            i += 1
            continue

        for lvl, pref in ((4, "#### "), (3, "### "), (2, "## "), (1, "# ")):
            if s.startswith(pref):
                add_runs(doc.add_heading(level=lvl), s[len(pref):])
                break
        else:
            if s.strip() in ("---", "***", "___"):
                pass
            elif re.match(r"^\s*[-*] ", s):
                add_runs(doc.add_paragraph(style="List Bullet"), re.sub(r"^\s*[-*] ", "", s))
            elif re.match(r"^\s*\d+\. ", s):
                add_runs(doc.add_paragraph(style="List Number"), re.sub(r"^\s*\d+\. ", "", s))
            elif s.strip():
                add_runs(doc.add_paragraph(), s)
        i += 1


doc = Document()
st = doc.styles["Normal"]
st.font.name = "Times New Roman"; st.font.size = Pt(11)
st.paragraph_format.space_after = Pt(4); st.paragraph_format.line_spacing = 1.15

doc.add_heading("АНВ Гарант: задание на правки сайта", level=0)
doc.add_paragraph("18 ноября 2026.")
add_runs(doc.add_paragraph(),
         "Готовые тексты страниц — в приложениях А–Е этого документа. "
         "Две статьи для /blog/ — отдельным пакетом файлов, в этот "
         "документ не входят: см. блок 7.")
doc.add_heading("Что в файле", level=1)
for line in ["Часть 1. Задание: что сломано, что править, чем проверять."] + [f[0] + "." for f in FILES]:
    add_runs(doc.add_paragraph(style="List Bullet"), line)
add_runs(doc.add_paragraph(),
         "**Стартовая точка:** пока не изменены title и H1, новые тексты эффекта "
         "не дадут. Порядок работ — в части 1, блоки 1–5. **Адреса страниц не "
         "меняем:** старые уже в индексе.")

doc.add_page_break()
p = doc.add_paragraph(); r = p.add_run("Часть 1. Задание"); r.bold = True; r.font.size = Pt(13)
render(read("docs/tz-vladimir_2026-11-18.md"))

for title, rel in FILES:
    doc.add_page_break()
    doc.add_heading(title, level=1)
    if "#1" in rel or "#2" in rel:
        path, part = rel.split("#")
        blocks = re.split(r"\n(?=# [12]\. )", read(path))
        md = [b for b in blocks if b.startswith("# " + part + ".")][0]
    else:
        md = read(rel)
    # Первый заголовок текста дублирует название приложения — убираем его,
    # чтобы в документе не было двух H1 подряд.
    md = re.sub(r"\A# [^\n]*\n+", "", md)
    render(md)

doc.save(OUT)
print("saved", OUT, os.path.getsize(OUT) // 1024, "KB")
