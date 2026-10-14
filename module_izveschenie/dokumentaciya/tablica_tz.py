# -*- coding: utf-8 -*-
"""Таблица технического задания для документа (раздел 5 ТЗ).

Состав столбцов заявки переменный, поэтому таблица собирается, а не задаётся
жёсткой формой: служебные столбцы убираются, столбец цены и типовые названия
приводятся к виду документации, остальные переносятся как есть (решения
специалиста от 07.10.2026, CLAUDE.md). Правила — списки ниже; добавляются
строками, без изменения кода.
"""
from __future__ import annotations

import re

import pravila

# служебные столбцы заявки, которые в документацию не идут (раздел 5 ТЗ):
# «Подразделение», «ОКВЭД», «КТРУ/ОКПД2» и столбец суммы — как бы он ни назывался
SLUZHEBNYE = (
    "подразделение", "нкми", "окпд2", "оквэд", "оквэд2",
    "ктру/окпд2", "кодокпд2", "тарифный", "торговая надбавка",
)
# сумма по позиции (количество x цена) — название всегда начинается с «сумма»
# или «общая сумма», как бы она дальше ни называлась («Сумма, руб.» и т.п.)
SLUZHEBNYE_PREFIKSY = ("источники финансирования", "источник финансирования", "фф омс", "канал финансирования",
                       "укб", "сумма", "общая сумма")

PEREIMENOVAT = {
    "наименования": "Наименование",
    "технические характеристики, упаковка": "Технические характеристики",
}
NAZVANIE_CENY = "Начальная максимальная цена за ед. измерения"
# форма заявки с листом «Расчет НМЦК» (426зэа, 333зэа): названия — как в
# документации 426зэа; узнаётся по столбцу наименования этой формы
PRIZNAK_FORMY_426 = "наименование, краткие характеристики, упаковка, производитель, страна происхождения"
PEREIMENOVAT_FORMA_426 = {
    PRIZNAK_FORMY_426: "Наименование, краткие характеристики, упаковка",
    "кол-во": "Кол-во ед. изм.",
    "цена за ед. изм., руб.": "Начальная максимальная Цена за ед. измер., руб. с НДС",
    "ставка ндс, %": "Применяемая ставка НДС, %",
    "срок годности не менее": "Срок годности, не менее",
    # шапка таблицы характеристик из отдельного файла ТЗ
    "ед. изм": "Ед. изм.",
}
# столбцы, где число печатается с копейками: «2 952,33», «22,00»
DENEZHNYE_FRAGMENTY = ("цена", "ставка ндс")


def pereimenovat_zagolovok(nazvanie: str, forma_426: bool = False) -> str:
    """Название столбца — к виду документации по списку замен (раздел 5 ТЗ);
    используется и при сборке таблицы из позиций заявки, и при переносе
    таблицы из отдельного файла ТЗ."""
    nizhnij = " ".join(nazvanie.lower().split())
    if forma_426 and nizhnij in PEREIMENOVAT_FORMA_426:
        return PEREIMENOVAT_FORMA_426[nizhnij]
    if nizhnij in PEREIMENOVAT:
        return PEREIMENOVAT[nizhnij]
    if nizhnij.startswith("цена за ед"):
        return NAZVANIE_CENY
    return nazvanie.replace("\n", "")


def chislo_s_kopejkami(tekst: str) -> str:
    """Число из ячейки заявки («2952.33», «22») — «2 952,33», «22,00» с
    неразрывным пробелом в разрядах; не число — как есть."""
    from decimal import Decimal

    from izvlechenie_excel import format_chislo
    stroka = str(tekst).strip()
    if re.fullmatch(r"-?\d+(?:[.,]\d+)?", stroka):
        return format_chislo(Decimal(stroka.replace(",", ".")))
    return tekst


def eto_forma_426(zagolovki: list[str]) -> bool:
    return any(" ".join(z.lower().split()) == PRIZNAK_FORMY_426 for z in zagolovki)


def pustoj_stolbec(pozicii: list[dict], nazvanie: str) -> bool:
    """Столбец пуст во всех строках — в документ не идёт."""
    return all(not str((p.get(nazvanie) or {}).get("znachenie") or "").strip() for p in pozicii)

# столбец меры по ПП № 1875 переносится, только если он есть в заявке
# (тогда последний, удалить нельзя) — решение специалиста от 09.10.2026
NAZVANIE_STOLBCA_1875 = "Постановление Правительства РФ от 23.12.2024 N 1875"

# фиксированные ширины столбцов (твипы), как в tz_iz_excel.py; остальные —
# поровну из оставшегося места
SHIRINY_STOLBCOV = {
    "№": 600,
    NAZVANIE_CENY: 1800,
    "НДС": 1200,
    "Ед. изм.": 1000,
    "Кол-во ед. изм.": 1000,
    "Остаточный срок годности": 1700,
}


def _sluzhebnyj(nazvanie: str) -> bool:
    nizhnij = nazvanie.lower().strip()
    # пробелы внутри названия не считаются — в заявках встречается то «КТРУ/ОКПД2»,
    # то «КТРУ/ ОКПД2»
    bez_probelov = re.sub(r"\s+", "", nizhnij)
    if bez_probelov in SLUZHEBNYE or nizhnij in SLUZHEBNYE:
        return True
    return nizhnij.startswith(SLUZHEBNYE_PREFIKSY)


def znachenie_stolbca_nds(tekst: str) -> str:
    """В таблице ТЗ столбец НДС печатает короткую форму — «не облагается» или
    ставку («10%»), без слова «НДС» и лишних слов исходной ячейки заявки
    (раздел 5 ТЗ: «в ячейке печатается «не облагается»»)."""
    nizhnij = tekst.lower()
    if "не облагается" in nizhnij:
        return "не облагается"
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*%", tekst)
    if m:
        return f"{m.group(1).replace('.', ',')}%"
    return tekst


def znachenie_stolbca_1875(tekst: str) -> str:
    """В таблице ТЗ столбец меры по ПП № 1875 печатает только слово —
    «Запрет», «Ограничение» или «Преимущество» (раздел 7 ТЗ), без номера
    позиции приложения и без скобок исходного текста заявки."""
    mera, _ = pravila.razobrat_stolbec_pp1875(tekst)
    slova = {"zapret": "Запрет", "ogranichenie": "Ограничение", "preimushchestvo": "Преимущество"}
    return slova.get(mera, tekst)


def sobrat_tablicu_tz(pozicii: list[dict]) -> tuple[list[str], list[list[str]]]:
    """pozicii — список словарей «название столбца заявки -> {znachenie,...}»
    (dokumentaciya.izvlechenie_excel). Возвращает (заголовки, строки)."""
    if not pozicii:
        return [], []

    stolbcy_zayavki = list(pozicii[0].keys())
    ostavit = [n for n in stolbcy_zayavki if not _sluzhebnyj(n) and not pustoj_stolbec(pozicii, n)]
    forma_426 = eto_forma_426(stolbcy_zayavki)

    zagolovki = []
    for nazvanie in ostavit:
        if "1875" in nazvanie:
            # столбец меры по ПП № 1875 — только если он есть в заявке,
            # тогда последний (раздел 5 ТЗ); название у разных заявок разное
            zagolovki.append(NAZVANIE_STOLBCA_1875)
        else:
            zagolovki.append(pereimenovat_zagolovok(nazvanie, forma_426))

    stroki = []
    for poziciya in pozicii:
        stroka = []
        for nazvanie in ostavit:
            znachenie = poziciya[nazvanie]["znachenie"]
            if "1875" in nazvanie:
                znachenie = znachenie_stolbca_1875(znachenie)
            elif nazvanie.lower().strip() == "ндс":
                znachenie = znachenie_stolbca_nds(znachenie)
            elif any(f in pereimenovat_zagolovok(nazvanie, forma_426).lower() for f in DENEZHNYE_FRAGMENTY):
                znachenie = chislo_s_kopejkami(znachenie)
            stroka.append(znachenie)
        stroki.append(stroka)
    return zagolovki, stroki


def dobavit_otsutstvuyushchie_stolbcy(
    zagolovki: list[str],
    stroki: list[list[str]],
    garantijnyj_srok: str | None = None,
    stavka_nds: str | None = None,
) -> tuple[list[str], list[list[str]], list[str]]:
    """Столбцы со сроком (гарантийным или годности) и НДС добавляются из
    условий заявки, если их нет в таблице ТЗ (раздел 5 ТЗ). Столбец срока —
    после характеристик; НДС — после цены, перед мерой по ПП № 1875, если
    она последняя. Возвращает (заголовки, строки, какие столбцы добавлены —
    для подписи источника «добавлен: взят из …»)."""
    zagolovki = list(zagolovki)
    stroki = [list(s) for s in stroki]
    dobavleno: list[str] = []
    nizhnie = [z.lower() for z in zagolovki]

    def vstavit_posle_fragmenta(fragment: str, nazvanie: str, znachenie: str, v_konec_esli_net: bool = True) -> None:
        idx = next((i for i, z in enumerate(nizhnie) if fragment in z), None)
        if idx is None:
            idx = len(zagolovki) - 1 if v_konec_esli_net else len(zagolovki)
        zagolovki.insert(idx + 1, nazvanie)
        nizhnie.insert(idx + 1, nazvanie.lower())
        for stroka in stroki:
            stroka.insert(idx + 1, znachenie)
        dobavleno.append(nazvanie)

    if garantijnyj_srok and not any("срок" in z for z in nizhnie):
        vstavit_posle_fragmenta("характеристик", "Гарантийный срок", garantijnyj_srok)

    if stavka_nds and not any(z == "ндс" for z in nizhnie):
        # перед столбцом меры по ПП № 1875, если он последний; иначе — после цены
        if zagolovki and "1875" in zagolovki[-1]:
            idx = len(zagolovki) - 2
            zagolovki.insert(idx + 1, "НДС")
            nizhnie.insert(idx + 1, "ндс")
            for stroka in stroki:
                stroka.insert(idx + 1, stavka_nds)
            dobavleno.append("НДС")
        else:
            vstavit_posle_fragmenta("цена", "НДС", stavka_nds)

    return zagolovki, stroki, dobavleno
