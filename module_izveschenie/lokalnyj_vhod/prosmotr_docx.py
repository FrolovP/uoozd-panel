# -*- coding: utf-8 -*-
"""Экран просмотра перед сохранением: документ Word (.docx) — в HTML для
браузера. Текст берётся из самого файла абзац за абзацем (в том же порядке,
что и в Word), поэтому просмотр совпадает с тем, что будет сохранено.
Значения из карточки подсвечиваются и правятся прямо на экране (data-pole —
поле карточки); остальной текст шаблона правится в Word после сохранения."""
from __future__ import annotations

import html
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

# значения переключателей и короткие слова не подсвечиваем — они встречаются
# в тексте шаблона и сами по себе ничего не говорят
_NE_PODSVECHIVAT = {"установлено", "не установлено", "применяется", "не применяется", "требуются",
                    "не требуются", "требуется", "не требуется", "да", "нет"}


def tekst_abzaca(p) -> str:
    """Текст абзаца так, как его видит Word: текст, табуляции, переносы строк."""
    kuski = []
    for el in p.iter():
        if el.tag == qn("w:t") and el.text:
            kuski.append(el.text)
        elif el.tag == qn("w:tab"):
            kuski.append("\t")
        elif el.tag in (qn("w:br"), qn("w:cr")):
            kuski.append("\n")
    return "".join(kuski)


def teksty_docx(put: str | Path) -> list[str]:
    """Тексты всех абзацев документа (включая таблицы) в порядке чтения."""
    return [tekst_abzaca(p) for p in Document(put).element.body.iter(qn("w:p"))]


def polya_dlya_podsvetki(sostoyanie: dict) -> list[tuple[str, str]]:
    """(значение, поле карточки) — что подсвечивать, длинные значения первыми."""
    znach = ((sostoyanie.get("karta") or {}).get("znacheniya") or {})
    para: list[tuple[str, str]] = []

    def dobavit(tekst, pole):
        if isinstance(tekst, str):
            tekst = tekst.strip()
            if len(tekst) >= 6 and tekst.lower() not in _NE_PODSVECHIVAT and not pole.startswith("_"):
                para.append((tekst, pole))

    for imya, z in znach.items():
        if imya in ("god",) or imya.startswith(("nr_", "dok_1875", "treb_", "iskl_")):
            continue
        if isinstance(z, str):
            dobavit(z, imya)
        elif isinstance(z, dict):
            for k, v in z.items():
                dobavit(v, f"{imya}.{k}")
        elif isinstance(z, list):
            for i, v in enumerate(z):
                if isinstance(v, dict):
                    for k, vv in v.items():
                        dobavit(vv, f"{imya}.{i}.{k}")
                else:
                    dobavit(v, f"{imya}[{i}]")  # один абзац списка
    if sostoyanie.get("nomer_zakupki"):
        para.append((str(sostoyanie["nomer_zakupki"]), "nomer_zakupki"))
    para.sort(key=lambda x: -len(x[0]))
    return para


def _s_podsvetkoj(tekst: str, podsvetka: list[tuple[str, str]]) -> str:
    """HTML абзаца: значения карточки — правимые вставки, без пересечений."""
    zanyato: list[tuple[int, int, str]] = []
    for znachenie, pole in podsvetka:
        for m in re.finditer(re.escape(znachenie), tekst):
            a, b = m.span()
            if any(a < kb and b > ka for ka, kb, _ in zanyato):
                continue
            zanyato.append((a, b, pole))
    zanyato.sort()
    kuski, poz = [], 0
    for a, b, pole in zanyato:
        kuski.append(html.escape(tekst[poz:a]))
        kuski.append(
            f'<span class="znach" data-pole="{html.escape(pole)}" tabindex="0" role="button" '
            f'title="Значение из карточки — щёлкните, чтобы исправить">{html.escape(tekst[a:b])}</span>'
        )
        poz = b
    kuski.append(html.escape(tekst[poz:]))
    return "".join(kuski).replace("\n", "<br>").replace("\t", "&emsp;")


def _abzac_html(p, podsvetka) -> str:
    tekst = tekst_abzaca(p)
    jc = p.find(f"{qn('w:pPr')}/{qn('w:jc')}")
    vyravnivanie = jc.get(qn("w:val")) if jc is not None else ""
    stil = {"center": "text-align:center", "right": "text-align:right", "both": "text-align:justify"}.get(vyravnivanie, "")
    runy = [r for r in p.iter(qn("w:r")) if any(t.text for t in r.iter(qn("w:t")))]
    zhirnyj = bool(runy) and all(r.find(f"{qn('w:rPr')}/{qn('w:b')}") is not None for r in runy)
    soderzhimoe = _s_podsvetkoj(tekst, podsvetka) or "&nbsp;"
    if zhirnyj:
        soderzhimoe = f"<b>{soderzhimoe}</b>"
    return f'<p style="{stil}">{soderzhimoe}</p>'


def _tablica_html(tbl, podsvetka) -> str:
    stroki = []
    for tr in tbl.iterchildren(qn("w:tr")):
        yacheyki = []
        for tc in tr.iterchildren(qn("w:tc")):
            span = tc.find(f"{qn('w:tcPr')}/{qn('w:gridSpan')}")
            colspan = f' colspan="{span.get(qn("w:val"))}"' if span is not None else ""
            yacheyki.append(f"<td{colspan}>{_blok_html(tc, podsvetka)}</td>")
        stroki.append(f"<tr>{''.join(yacheyki)}</tr>")
    return f'<table class="doc-t">{"".join(stroki)}</table>'


def _blok_html(element, podsvetka) -> str:
    kuski = []
    for el in element.iterchildren():
        if el.tag == qn("w:p"):
            kuski.append(_abzac_html(el, podsvetka))
        elif el.tag == qn("w:tbl"):
            kuski.append(_tablica_html(el, podsvetka))
        elif el.tag == qn("w:sdt"):
            soderzhimoe = el.find(qn("w:sdtContent"))
            if soderzhimoe is not None:
                kuski.append(_blok_html(soderzhimoe, podsvetka))
    return "".join(kuski)


def docx_v_html(put: str | Path, podsvetka: list[tuple[str, str]]) -> str:
    """Весь документ — HTML с правимыми значениями карточки."""
    return _blok_html(Document(put).element.body, podsvetka)


def tekst_iz_html(html_tekst: str) -> str:
    """Что видит сотрудник на экране просмотра — для сверки с файлом."""
    bez_tegov = re.sub(r"<br>", "\n", html_tekst)
    bez_tegov = re.sub(r"</p>|</td>", "\n", bez_tegov)
    bez_tegov = re.sub(r"<[^>]+>", "", bez_tegov)
    return html.unescape(bez_tegov.replace("&emsp;", "\t").replace("&nbsp;", ""))
