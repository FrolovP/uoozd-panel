# -*- coding: utf-8 -*-
"""Самопроверка карточки (раздел 11 ТЗ). Пока здесь — только проверка
возможных опечаток в инструкции по описанию товара (решение специалиста,
раздел «Главные ограничения» CLAUDE.md): инструкция переносится дословно,
программа ничего не правит сама, а только предупреждает.
"""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

import raschety

_SLOVO = re.compile(r"[а-яёА-ЯЁ]+")


@dataclass
class VozmozhnayaOpechatka:
    """Предупреждение «возможная опечатка: найдено → исправить» (кнопка
    «Исправить» в карточке; молча ничего не меняется)."""

    najdeno: str
    ispravlenie: str

    def __str__(self) -> str:
        return f"возможная опечатка: {self.najdeno} → {self.ispravlenie}"


def zagruzit_slovar(put: str | Path) -> set[str]:
    """«Словарь для проверки опечаток.xlsx» — первый столбец листа с данными
    (не «Как заполнять»), без заголовка."""
    wb = openpyxl.load_workbook(put, data_only=True)
    ws = wb.worksheets[0]
    return {
        str(row[0]).strip().lower() for row in ws.iter_rows(min_row=2, max_col=1, values_only=True)
        if row[0]
    }


def _odna_pravka_ne_v_okonchanii(a: str, b: str) -> bool:
    """Слова отличаются ровно одной буквой (лишняя, пропущенная или другая),
    и не в последних двух буквах: «ограничение / ограничения» — другая форма
    слова, а не опечатка; «применятся / применяется» — опечатка."""
    if a == b or abs(len(a) - len(b)) > 1:
        return False
    i = 0
    while i < min(len(a), len(b)) and a[i] == b[i]:
        i += 1
    if len(a) == len(b):
        # другая буква в окончании («применяются / применяется») — форма слова
        return a[i + 1:] == b[i + 1:] and i < len(a) - 4
    if len(a) > len(b):
        sovpadaet = a[i + 1:] == b[i:]
    else:
        sovpadaet = a[i:] == b[i + 1:]
    # пропущенная или лишняя буква («применятся», «сиволов») — опечатка
    return sovpadaet and i < min(len(a), len(b)) - 2


def _s_registrom(obrazec: str, slovo: str) -> str:
    return slovo[:1].upper() + slovo[1:] if obrazec[:1].isupper() else slovo


def naiti_vozmozhnye_opechatki(abzacy: list[str], slovar: set[str]) -> list[VozmozhnayaOpechatka]:
    """Возможные опечатки (слова длиннее пяти букв):
    - окончание «-иого», которого в русском языке не бывает → «-ого»
      («монокристаллическиого» → «монокристаллического»), словарь не нужен;
    - слова нет в словаре, а словарное отличается от него одной буквой не в
      окончании («сиволов» → «символов», «применятся» → «применяется»).
    Короткие слова не проверяются: случайных совпадений слишком много."""
    opechatki: list[VozmozhnayaOpechatka] = []
    uzhe_naideno: set[str] = set()
    for abzac in abzacy:
        for sovpadenie in _SLOVO.finditer(abzac):
            slovo = sovpadenie.group()
            ochishchennoe = slovo.lower()
            if len(ochishchennoe) < 6 or ochishchennoe in slovar or ochishchennoe in uzhe_naideno:
                continue
            if ochishchennoe.endswith("иого"):
                opechatki.append(VozmozhnayaOpechatka(slovo, _s_registrom(slovo, ochishchennoe[:-4] + "ого")))
                uzhe_naideno.add(ochishchennoe)
                continue
            pohozhie = [
                p for p in difflib.get_close_matches(ochishchennoe, slovar, n=5, cutoff=0.82)
                if _odna_pravka_ne_v_okonchanii(ochishchennoe, p)
            ]
            if pohozhie:
                opechatki.append(VozmozhnayaOpechatka(slovo, _s_registrom(slovo, pohozhie[0])))
                uzhe_naideno.add(ochishchennoe)
    return opechatki


@dataclass
class RezultatSamoproverki:
    """Раздел 11 ТЗ: ошибки не дают сформировать документ, предупреждения —
    не останавливают расчёт."""

    oshibki: list[str] = field(default_factory=list)
    preduprezhdeniya: list[str] = field(default_factory=list)

    @property
    def mozhno_sformirovat(self) -> bool:
        return not self.oshibki


def proverit_kartochku(karta, slovar_instrukcii: set[str] | None = None) -> RezultatSamoproverki:
    """Собирает проверки разделов 6, 7 и 9 ТЗ в одну самопроверку карточки,
    извлечённой из заявки Excel."""
    oshibki: list[str] = []
    preduprezhdeniya: list[str] = list(karta.preduprezhdeniya)

    if slovar_instrukcii and karta.znacheniya.get("instrukciya_tovara"):
        opechatki = naiti_vozmozhnye_opechatki(karta.znacheniya["instrukciya_tovara"], slovar_instrukcii)
        preduprezhdeniya.extend(str(o) for o in opechatki)

    nmcd_tekst = karta.znacheniya.get("nmcd")
    if karta.pozicii and nmcd_tekst:
        oshibki.extend(raschety.proverit_nmcd_po_poziciyam(karta.pozicii, raschety.chislo(nmcd_tekst)))

    nmck_raschet = karta.znacheniya.get("nmck_raschet") or {}
    if nmck_raschet.get("kp"):
        ceny = [raschety.chislo(k["summa"]) for k in nmck_raschet["kp"]]
        pereschet = raschety.rasschitat_nmck_po_kp(ceny)
        preduprezhdeniya.extend(pereschet.preduprezhdeniya)
        for pole, znachenie in (
            ("srednyaya", pereschet.srednyaya), ("sko", pereschet.sko), ("koef_var", pereschet.koef_var),
        ):
            if pole in nmck_raschet and not raschety.sverit_s_dopuskom(raschety.chislo(nmck_raschet[pole]), znachenie):
                oshibki.append(f"nmck_raschet.{pole}: в заявке {nmck_raschet[pole]}, при пересчёте {znachenie}")

    if "_nmck_tarif" in karta.znacheniya:
        oshibki.extend(raschety.proverit_tarifnyj_raschet(karta.znacheniya["_nmck_tarif"]))

    if nmcd_tekst and karta.pozicii:
        stavka = raschety.izvlech_stavku_nds(karta.pozicii)
        summa_v_zayavke = raschety.izvlech_summu_iz_nds_teksta(karta.znacheniya.get("nds_tekst", ""))
        if stavka is not None and summa_v_zayavke is not None:
            rasschitannaya = raschety.summa_nds(raschety.chislo(nmcd_tekst), stavka)
            if not raschety.sverit_s_dopuskom(summa_v_zayavke, rasschitannaya, raschety.DOPUSK_NDS):
                preduprezhdeniya.append(
                    f"сумма НДС в заявке {summa_v_zayavke}, при пересчёте {rasschitannaya} "
                    f"(допуск {raschety.DOPUSK_NDS} руб.)"
                )

    return RezultatSamoproverki(oshibki=oshibki, preduprezhdeniya=preduprezhdeniya)
