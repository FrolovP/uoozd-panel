# -*- coding: utf-8 -*-
"""Переключатели нацрежима и требований (раздел 7 ТЗ).

Правила определяют мeру по ПП № 1875 из текста заявки («Иные условия
исполнения договора») и расставляют переключатели пп. 5 и 17. Набор
девяти требований п. 15 и признак п. 14 — значения по умолчанию для типовой
поставки товара (правило 11 раздела 7 ТЗ); меняются файлом правил, не кодом.
"""
from __future__ import annotations

import re

# правило 11 (раздел 7 ТЗ): набор по умолчанию для типовой поставки товара
TREBOVANIYA_P15_PO_UMOLCHANIYU = {
    "treb_rnp223": "требуется",
    "treb_rnp44": "требуется",
    "treb_opyt": "не требуется",
    "treb_resursy": "не требуется",
    "treb_harakteristiki": "требуется",
    "treb_uklonenie": "требуется",
    "treb_uchrediteli": "требуется",
    "treb_spravka_fns": "не требуется",
    "treb_cepochka": "не требуется",
}
ISKL_PRAVA_PO_UMOLCHANIYU = "не требуется"


def opredelit_meru(nacrezhim_tekst: str) -> str | None:
    """«запрет» / «ограничение» / «преимущество» — по тексту заявки, или None,
    если постановление к лоту не применяется."""
    nizhnij = nacrezhim_tekst.lower()
    if "не применяется" in nizhnij:
        return None
    if "преимущество" in nizhnij:
        return "preimushchestvo"
    if "ограничение" in nizhnij:
        return "ogranichenie"
    if "запрет" in nizhnij:
        return "zapret"
    return None


def opredelit_blok_1875(nacrezhim_tekst: str, mera: str | None) -> int | None:
    """Какой из пяти блоков п. 17 (раздел 7 ТЗ) относится к мере:
    1 — запрет, Приложение №1, поз. 1–145; 2 — запрет, поз. 146 (ПО);
    3 — ограничение, Приложение №2, поз. 1–432; 4 — ограничение, поз. 433
    (лекарства); 5 — преимущество (вне приложений)."""
    if mera == "preimushchestvo":
        return 5
    if mera is None:
        return None
    m = re.search(r"\(п\.\s*(\d+)", nacrezhim_tekst)
    poz_v_prilozhenii = int(m.group(1)) if m else None
    if mera == "zapret":
        return 2 if poz_v_prilozhenii == 146 else 1
    if mera == "ogranichenie":
        return 4 if poz_v_prilozhenii == 433 else 3
    return None


def rasschitat_nacrezhim(nacrezhim_tekst: str) -> dict[str, str]:
    """Переключатели п. 5 и дополняющие их блоки п. 17 (без номеров позиций:
    правило для лота с одной мерой на все позиции — раздел 7 ТЗ)."""
    mera = opredelit_meru(nacrezhim_tekst)
    blok = opredelit_blok_1875(nacrezhim_tekst, mera)

    rezultat = {
        "nr_primenyaetsya": "Применяется" if mera else "Не применяется",
        "nr_zapret": "Установлено" if mera == "zapret" else "Не установлено",
        "nr_ogranichenie": "Установлено" if mera == "ogranichenie" else "Не установлено",
        "nr_preimushchestvo": "Установлено" if mera == "preimushchestvo" else "Не установлено",
    }
    for n in range(1, 6):
        rezultat[f"dok_1875_{n}"] = "требуются" if blok == n else "не требуются"
    return rezultat


def razobrat_stolbec_pp1875(tekst: str) -> tuple[str | None, int | None]:
    """Столбец «Постановление ... N 1875» в таблице позиций Word-заявки
    (раздел 5 ТЗ, пример 1500/а-26): «Приложение № 2 (п. 200)» -> ограничение,
    «Приложение №1 (п. ...)» -> запрет, «Отсутствует в приложении №1,2» ->
    преимущество."""
    t = tekst.strip()
    if not t:
        return None, None
    if "отсутствует" in t.lower():
        return "preimushchestvo", None
    m = re.search(r"Приложени[ея]\s*№?\s*(\d)\s*\(п\.\s*(\d+)\)", t)
    # форма 426зэа/333зэа: «Перечень 2 (13), применяется ограничение»
    m = m or re.search(r"Переч[её]нь\s*№?\s*(\d)\s*\((?:п\.\s*)?(\d+)\)", t)
    if not m:
        return None, None
    prilozhenie, poz_v_prilozhenii = int(m.group(1)), int(m.group(2))
    return ("zapret" if prilozhenie == 1 else "ogranichenie"), poz_v_prilozhenii


def _blok_1875(mera: str | None, poz_v_prilozhenii: int | None) -> int | None:
    if mera == "preimushchestvo":
        return 5
    if mera is None:
        return None
    if mera == "zapret":
        return 2 if poz_v_prilozhenii == 146 else 1
    if mera == "ogranichenie":
        return 4 if poz_v_prilozhenii == 433 else 3
    return None


def rasschitat_nacrezhim_po_poziciyam(mery_poziciy: list[tuple[int, str | None, int | None]]) -> dict[str, str]:
    """Переключатели п. 5 и блоки п. 17 для лота, где у позиций разные меры
    (раздел 7 ТЗ, пример 1500/а-26: ограничение — позиции 1–3, преимущество —
    позиция 4). mery_poziciy — [(номер позиции, мера, позиция в приложении), ...].
    Если мера одна на весь лот — номера позиций не печатаются (решение
    специалиста от 09.10.2026); если разные — печатаются у каждой меры."""
    vsego = len(mery_poziciy)
    poz_po_mere: dict[str, list[int]] = {"zapret": [], "ogranichenie": [], "preimushchestvo": []}
    poz_po_bloku: dict[int, list[int]] = {n: [] for n in range(1, 6)}
    for nomer, mera, poz_v_prilozhenii in mery_poziciy:
        if mera in poz_po_mere:
            poz_po_mere[mera].append(nomer)
        blok = _blok_1875(mera, poz_v_prilozhenii)
        if blok:
            poz_po_bloku[blok].append(nomer)

    def s_nomerami(pozicii: list[int]) -> str:
        if not pozicii:
            return ""
        if len(pozicii) == vsego:
            return "Установлено"
        slovo = "позиция" if len(pozicii) == 1 else "позиции"
        return f"Установлено ({slovo} {', '.join(map(str, pozicii))} Технического задания)"

    primenyaetsya = any(poz_po_mere.values())
    rezultat = {"nr_primenyaetsya": "Применяется" if primenyaetsya else "Не применяется"}
    for mera, imya in (("zapret", "nr_zapret"), ("ogranichenie", "nr_ogranichenie"), ("preimushchestvo", "nr_preimushchestvo")):
        rezultat[imya] = s_nomerami(poz_po_mere[mera]) or "Не установлено"

    for n in range(1, 6):
        pozicii = poz_po_bloku[n]
        if not pozicii:
            rezultat[f"dok_1875_{n}"] = "не требуются"
        elif len(pozicii) == vsego:
            rezultat[f"dok_1875_{n}"] = "требуются"
        else:
            slovo = "позиция" if len(pozicii) == 1 else "позиции"
            rezultat[f"dok_1875_{n}"] = f"требуются — {slovo} {', '.join(map(str, pozicii))} Технического задания"
    return rezultat


def _normalizovat(tekst: str) -> str:
    """Для сравнения с эталонным списком требований: регистр, пробелы и
    дефисы не считаются — у разных заявок одна и та же стандартная строка
    написана чуть по-разному («223-ФЗ» / «223ФЗ»)."""
    return re.sub(r"[\s\-.,;:]+", "", tekst.lower())


def sovpadaet_so_standartnym(stroka: str, standartnye: set[str]) -> bool:
    """Стандартная строка, в том числе без пояснения в скобках в конце
    («Отсутствие в реестре недобросовестных поставщиков (223-ФЗ)» = «…поставщиков.»)."""
    normalizovannaya = _normalizovat(stroka)
    bez_skobok = _normalizovat(re.sub(r"\s*\([^)]*\)\s*\.?\s*$", "", stroka))
    for s in standartnye:
        etalon = _normalizovat(s)
        etalon_bez_skobok = _normalizovat(re.sub(r"\s*\([^)]*\)\s*\.?\s*$", "", s))
        if normalizovannaya in (etalon, etalon_bez_skobok) or bez_skobok == etalon_bez_skobok:
            return True
    return False


def opredelit_ru_minzdrav(trebovaniya_k_uchastnikam: list[str]) -> str:
    """Регистрационные удостоверения (п. 17, раздел 4 ТЗ) — требуются, если
    заявка явно требует их от участника."""
    # «Регистрационные удостоверения — не требуются» — не требуются
    trebuetsya = any(
        "регистрационные удостоверения" in t.lower() and not re.search(r"\bне\s+требу", t, re.I)
        for t in trebovaniya_k_uchastnikam
    )
    return "требуются" if trebuetsya else "не требуются"
