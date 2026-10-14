# -*- coding: utf-8 -*-
"""Дополняет карточку, извлечённую из заявки, переключателями и расчётами
разделов 6 и 7 ТЗ (этап 2). Источник добавленных полей — «правила» или
«рассчитано»; извлечение из заявки (этап 1) не трогает.
"""
from __future__ import annotations

from datetime import datetime

import pravila
import raschety
from izvlechenie_excel import Istochnik, Kartochka, format_chislo


def dopolnit(karta: Kartochka) -> None:
    _dopolnit_god(karta)
    _dopolnit_nacrezhim(karta)
    _dopolnit_trebovaniya_po_umolchaniyu(karta)
    _dopolnit_obesp_zayavki(karta)
    _dopolnit_obesp_ispolneniya(karta)


def _dopolnit_god(karta: Kartochka) -> None:
    """Год в строке «г. Москва, … год» (титульный лист) — текущий год, виден
    в карточке сразу, не только после «Сформировать» (раздел 4 ТЗ)."""
    if not karta.znacheniya.get("god"):
        _polozhit_pravilom(karta, "god", str(datetime.now().year), "рассчитано: текущий год")


def _polozhit_pravilom(karta: Kartochka, imya: str, znachenie, citata: str) -> None:
    karta.znacheniya[imya] = znachenie
    karta.istochniki[imya] = Istochnik("pravila", "", "", citata)


def _dopolnit_nacrezhim(karta: Kartochka) -> None:
    nacrezhim_tekst = karta.znacheniya.get("nacrezhim_tekst", "")
    for imya, znachenie in pravila.rasschitat_nacrezhim(nacrezhim_tekst).items():
        _polozhit_pravilom(karta, imya, znachenie, nacrezhim_tekst or "в заявке нацрежим не упомянут")
    # исключения из запрета (п. 17) — пустые, если заявка их не называет
    for imya in ("iskl_zapret_1", "iskl_zapret_2"):
        if imya not in karta.znacheniya:
            _polozhit_pravilom(karta, imya, "", "по умолчанию: исключений из запрета нет")

    ru_minzdrav = pravila.opredelit_ru_minzdrav(karta.znacheniya.get("trebovaniya_k_uchastnikam", []))
    _polozhit_pravilom(karta, "ru_minzdrav", ru_minzdrav, "список требований к участникам заявки")


def _dopolnit_trebovaniya_po_umolchaniyu(karta: Kartochka) -> None:
    for imya, znachenie in pravila.TREBOVANIYA_P15_PO_UMOLCHANIYU.items():
        _polozhit_pravilom(karta, imya, znachenie, "правило 11 (раздел 7 ТЗ): типовая поставка товара")
    _polozhit_pravilom(karta, "iskl_prava", pravila.ISKL_PRAVA_PO_UMOLCHANIYU, "значение по умолчанию")


def _dopolnit_obesp_zayavki(karta: Kartochka) -> None:
    nmcd = karta.znacheniya.get("nmcd")
    obesp = karta.znacheniya.get("obesp_zayavki")
    if not nmcd or obesp is None:
        return
    ustanovleno, preduprezhdeniya = raschety.primenit_pravilo_obespecheniya_zayavki(
        raschety.chislo(nmcd), obesp.get("ustanovleno", False), obesp.get("procent", ""),
    )
    if ustanovleno != obesp.get("ustanovleno"):
        karta.znacheniya["obesp_zayavki"] = {**obesp, "ustanovleno": ustanovleno}
        karta.istochniki["obesp_zayavki"] = Istochnik("pravila", "", "", "правило 8 (раздел 7 ТЗ): граница 5 000 000,00 руб.")
    karta.preduprezhdeniya.extend(preduprezhdeniya)


def _dopolnit_obesp_ispolneniya(karta: Kartochka) -> None:
    procent = karta.znacheniya.get("obesp_isp_procent")
    if procent is None:
        # заявка не упоминала обеспечение исполнения вовсе — 5% по умолчанию
        # (раздел 4 ТЗ); явное «не применяется» из заявки (procent == "") не трогаем
        procent = "5"
        _polozhit_pravilom(karta, "obesp_isp_procent", procent, "по умолчанию 5% (раздел 4 ТЗ)")

    if not procent:
        _polozhit_pravilom(karta, "obesp_isp_summa", "", "обеспечение исполнения договора не применяется")
        return

    nmcd_dostupen = karta.znacheniya.get("nmcd")
    if not nmcd_dostupen:
        return
    summa = raschety.summa_obespecheniya(raschety.chislo(nmcd_dostupen), raschety.chislo(procent))
    _polozhit_pravilom(
        karta, "obesp_isp_summa", format_chislo(summa),
        f"рассчитано: НМЦД {nmcd_dostupen} x {procent}% / 100",
    )
