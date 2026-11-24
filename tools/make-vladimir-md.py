#!/usr/bin/env python3
"""Собирает одно ТЗ в markdown: задание + готовые тексты восьми страниц.

Запуск:
  python3 tools/make-vladimir-md.py

Зачем скрипт: Владимиру уходит один файл, а не восемь. Markdown удобнее docx,
когда задание пересылают в мессенджере или открывают в редакторе. Источник
правды — docs/tz-vladimir_2026-11-18.md и docs/text-*.md; править сборку руками
нельзя, при пересборке правки исчезнут.
"""
import os
import re

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "docs", "Задание-по-сайту-2026-11-22.md")

TZ = "docs/tz-vladimir_2026-11-18.md"

FILES = [
    ("Приложение А. Страница «Проверка объекта перед покупкой»",
     "docs/text-proverka-obekta_2026-11-18.md"),
    ("Приложение Б. Страница «Юристы»", "docs/text-yuristy_2026-11-18.md"),
    ("Приложение В. Страница «Сделки с недвижимостью»",
     "docs/text-sdelki_2026-11-18.md"),
    ("Приложение Г. Страница «Анализ коммерческой недвижимости»",
     "docs/text-kommercia_2026-11-18.md#1"),
    ("Приложение Д. Страница «Продажа коммерческой недвижимости»",
     "docs/text-kommercia_2026-11-18.md#2"),
    ("Приложение Е. Страница «Запуск агентства недвижимости»",
     "docs/text-zapusk-agentstva_2026-11-18.md"),
    ("Приложение Ж. Страница «Геленджик»",
     "docs/text-gelendzhik_2026-11-22.md"),
]

HEADER = """# АНВ Гарант: задание на правки сайта

18 ноября 2026, редакция 22 ноября 2026.

Задание на правки `gk-anvgarant.ru`: что сломано, что править, готовые тексты
восьми страниц и команды приёмки. Файл самодостаточный — приложения внутри,
открывать соседние документы не нужно.

**Стартовая точка:** пока не изменены title и H1, новые тексты эффекта не дадут.
Порядок работ — часть 1, блоки 1–5.1. **Адреса страниц не меняем:** старые уже
в индексе.

Ссылки на внутренние документы убраны: всё, что нужно для работы, есть в этом
файле.

**Пометки `[H1]` и `[H2]` в приложениях** — это теги страницы, а не уровни
заголовков. Первый заголовок приложения — это `h1`, второй раз его не выводить;
в блоке «Мета» он стоит только для справки.

## Что в файле

- Часть 1. Задание: что сломано, что править, чем проверять.
{apps}

---

# Часть 1. Задание
"""


def read(rel):
    with open(os.path.join(BASE, rel), encoding="utf-8") as f:
        return f.read()


def plain_links(text):
    """Относительные ссылки на файлы репозитория в отправляемом файле не работают:
    оставляем только подпись. Внешние адреса на http(s) не трогаем."""
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", text)
    return re.sub(r"\[([^\]]+)\]\((?!https?:)[^)]*\)", r"\1", text)


def strip_first_h1(md):
    """Убирает первый заголовок первого уровня: он дублирует название части."""
    return re.sub(r"\A# [^\n]*\n+", "", md).lstrip("\n")


def appendix_md(rel):
    if "#" in rel:
        path, part = rel.split("#")
        blocks = re.split(r"\n(?=# [12]\. )", read(path))
        md = [b for b in blocks if b.startswith("# " + part + ".")][0]
    else:
        md = read(rel)
    return strip_first_h1(md).rstrip() + "\n"


def main():
    apps = "\n".join(f"- {title}" for title, _ in FILES)
    chunks = [HEADER.format(apps=apps), strip_first_h1(read(TZ)).rstrip(), ""]

    for title, rel in FILES:
        chunks.append("\n---\n")
        chunks.append(f"# {title}\n")
        chunks.append(appendix_md(rel))

    text = "\n".join(chunks)
    text = plain_links(text)
    text = re.sub(r"\n{3,}", "\n\n", text).rstrip() + "\n"
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(text)
    print("saved", OUT, len(text.split()), "слов")


if __name__ == "__main__":
    main()
