#!/usr/bin/env python3
"""Собирает читабельный docx со статьями для владельца компании (Нвера).

Запуск:
  python3 tools/make-nver-docx.py

Это не техническое задание: мета, разметка и служебные блоки убраны, остаются
заголовок, короткое «зачем эта статья» и сам текст. Источник правды —
docs/articles/*.md; править docx руками нельзя.
"""
import os
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "docs", "Статьи-на-прочтение-Нверу_2026-11-22.docx")

ARTICLES = [
    (
        "docs/articles/2026-11-18-deal-support.md",
        "Человек сомневается, нужен ли юрист, если риэлтор уже есть. Статья "
        "показывает работу по этапам и честно говорит, кому сопровождение не "
        "нужно. Ведёт на страницу «Сопровождение сделки».",
    ),
    (
        "docs/articles/2026-11-18-mortgage-deal.md",
        "Многие думают, что в ипотеке объект уже проверил банк. Статья разводит "
        "проверки банка, покупателя и продавца и объясняет, почему одобрение "
        "кредита не равно проверенному объекту. Ведёт туда же.",
    ),
    (
        "docs/articles/2026-11-18-dkp-five-places.md",
        "Договор подписывают, не читая, потому что «шаблон». Статья показывает "
        "пять мест, где написанное расходится с договорённостью. Инструкции, "
        "как проверить самому, в ней намеренно нет. Ведёт на страницу "
        "сопровождения.",
    ),
]

INTRO = [
    ("Три статьи для сайта", "title"),
    ("На прочтение · 22 ноября 2026", "subtitle"),
    (
        "Это тексты, которые появятся в разделе «Полезное» на сайте. Пишем их не "
        "для поиска, а чтобы человек, который выбирает между «позвонить» и "
        "«разобраться самому», увидел, как мы работаем. Рекламы внутри нет: "
        "только объяснение, за что платят.",
        "text",
    ),
    ("Зачем это нужно", "h2"),
    ("Люди перед сделкой ищут объяснение, а не рекламу. Эти три статьи отвечают "
     "на вопросы, которые чаще всего задают по телефону, — и человек приходит "
     "на звонок уже с пониманием, о чём разговор.", "bullet"),
    ("Каждая статья ведёт на страницу услуги: прочитал — понял, зачем нужен "
     "юрист, и куда нажать.", "bullet"),
    ("Тексты сверены с фактами компании, но часть формулировок — про то, как "
     "устроена ваша работа. Их лучше прочитать тому, кто эту работу делает.", "bullet"),
    ("Что просим сделать", "h2"),
    ("Прочитать — это 15–20 минут. Отметить, где написано неточно, слишком "
     "смело или не так, как вы работаете. Правки можно прямо в файле или "
     "списком.", "bullet"),
    ("Особенно первую статью: там про задаток, срок регистрации и состав "
     "проверки. Остальные две можно смотреть по желанию.", "bullet"),
    ("Чего в статьях нет намеренно: советов «сделайте так» по конкретному делу, "
     "пошаговых инструкций и обещаний результата. Статья не заменяет "
     "консультацию — это принципиально.", "bullet"),
]


def clean_inline(text):
    text = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", text)
    return text.replace("**", "")


def add_bold_runs(paragraph, text):
    for i, part in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if not part:
            continue
        run = paragraph.add_run(part)
        run.bold = i % 2 == 1


def article_body(path):
    """Возвращает (H1, [(тип, текст)]) без мета и служебных блоков."""
    s = open(os.path.join(BASE, path), encoding="utf-8").read()
    h1 = re.search(r"\*\*H1:\*\*\s*`([^`]+)`", s).group(1)
    body = re.search(r"## Текст\n(.*?)(?=\n## |\Z)", s, re.S).group(1)

    items, buf = [], []

    def flush():
        if not buf:
            return
        lines = [re.sub(r"^>\s?", "", l).rstrip() for l in buf]
        if lines and lines[0].startswith("- "):
            cur = None
            for l in lines:
                if l.startswith("- "):
                    if cur:
                        items.append(("li", clean_inline(cur.strip())))
                    cur = l[2:]
                elif l.startswith("  ") and cur is not None:
                    cur += " " + l.strip()
                elif cur is not None:
                    cur += " " + l.strip()
            if cur:
                items.append(("li", clean_inline(cur.strip())))
        else:
            items.append(("p", clean_inline(" ".join(l.strip() for l in lines))))
        buf.clear()

    for line in body.split("\n"):
        if line.startswith(">"):
            if "**[H2]" in line:
                flush()
                title = re.match(r">\s*\*\*\[H2\]\s*(.+?)\*\*", line).group(1)
                items.append(("h2", clean_inline(title)))
                continue
            content = re.sub(r"^>\s?", "", line).rstrip()
            if not content:
                flush()
            else:
                buf.append(line)
        else:
            flush()
    flush()
    return h1, items


def main():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.25

    for name, size, before in (("Heading 1", 18, 0), ("Heading 2", 13, 12)):
        st = doc.styles[name]
        st.font.name = "Calibri"
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0x11, 0x11, 0x11)
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(4)
        st.paragraph_format.keep_with_next = True

    for text, kind in INTRO:
        if kind == "title":
            p = doc.add_paragraph()
            r = p.add_run(text)
            r.bold = True
            r.font.size = Pt(24)
            p.paragraph_format.space_after = Pt(2)
        elif kind == "subtitle":
            p = doc.add_paragraph()
            r = p.add_run(text)
            r.font.size = Pt(11)
            r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
            p.paragraph_format.space_after = Pt(18)
        elif kind == "h2":
            p = doc.add_paragraph()
            r = p.add_run(text)
            r.bold = True
            r.font.size = Pt(13)
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)
        elif kind == "bullet":
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(4)
            add_bold_runs(p, text)
        else:
            add_bold_runs(doc.add_paragraph(), text)

    for path, why in ARTICLES:
        doc.add_page_break()
        h1, items = article_body(path)

        doc.add_heading(h1, level=1)

        p = doc.add_paragraph()
        r = p.add_run("Зачем эта статья")
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        p.paragraph_format.space_after = Pt(2)

        p = doc.add_paragraph()
        add_bold_runs(p, why)
        for run in p.runs:
            run.italic = True
            run.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
        p.paragraph_format.space_after = Pt(12)

        for kind, text in items:
            if kind == "h2":
                doc.add_heading(text, level=2)
            elif kind == "li":
                p = doc.add_paragraph(style="List Bullet")
                p.paragraph_format.space_after = Pt(2)
                add_bold_runs(p, text)
            else:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                add_bold_runs(p, text)

    doc.save(OUT)
    print("saved", OUT, os.path.getsize(OUT) // 1024, "KB")


if __name__ == "__main__":
    main()
