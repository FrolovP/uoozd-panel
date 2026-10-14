# -*- coding: utf-8 -*-
"""Расчёты (раздел 6 ТЗ): НМЦК, НДС, обеспечение. Деньги считаются Decimal,
не float. Округление — до копеек, половина округляется вверх.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

KOPEJKA = Decimal("0.01")
DOPUSK_NDS = Decimal("0.01")


def chislo(tekst: str | Decimal | int | float) -> Decimal:
    """Строка заявки («13 712,07», с обычным или неразрывным пробелом) — в Decimal."""
    if isinstance(tekst, Decimal):
        return tekst
    s = str(tekst).replace(" ", "").replace(" ", "").replace(",", ".")
    return Decimal(s)


def okruglit(d: Decimal) -> Decimal:
    return d.quantize(KOPEJKA, rounding=ROUND_HALF_UP)


@dataclass
class RaschetNmckPoKp:
    """Расчёт НМЦК по ценам нескольких КП (раздел 6 ТЗ): средняя, отклонение,
    коэффициент вариации — от неокруглённых значений, округляются только при
    выводе; однородность — по коэффициенту вариации."""

    srednyaya: Decimal
    sko: Decimal
    koef_var: Decimal
    odnorodnost: str
    preduprezhdeniya: list[str]


def rasschitat_nmck_po_kp(ceny: list[Decimal]) -> RaschetNmckPoKp:
    n = len(ceny)
    srednyaya_neokr = sum(ceny) / n
    if n > 1:
        variaciya = sum((c - srednyaya_neokr) ** 2 for c in ceny) / (n - 1)
        sko = variaciya.sqrt()
    else:
        sko = Decimal(0)
    koef_var = (sko / srednyaya_neokr * 100) if srednyaya_neokr else Decimal(0)

    preduprezhdeniya = []
    if n < 3:
        preduprezhdeniya.append(f"меньше трёх цен по позиции ({n})")
    if koef_var > 33:
        preduprezhdeniya.append(f"коэффициент вариации выше 33% ({okruglit(koef_var)})")
    odnorodnost = "ОДНОРОДНЫЕ" if koef_var <= 33 else "НЕОДНОРОДНЫЕ"

    return RaschetNmckPoKp(
        srednyaya=okruglit(srednyaya_neokr),
        sko=okruglit(sko),
        koef_var=okruglit(koef_var),
        odnorodnost=odnorodnost,
        preduprezhdeniya=preduprezhdeniya,
    )


def summa_nds(nmcd: Decimal, stavka_procent: Decimal) -> Decimal:
    """НДС «в том числе»: НМЦД x ставка / (100 + ставка) (раздел 6 ТЗ)."""
    return okruglit(nmcd * stavka_procent / (100 + stavka_procent))


def summa_obespecheniya(nmcd: Decimal, procent: Decimal) -> Decimal:
    """Обеспечение исполнения договора: НМЦД x процент / 100 (раздел 6 ТЗ)."""
    return okruglit(nmcd * procent / 100)


NMCD_GRANICA_OBESP_ZAYAVKI = Decimal("5000000.00")


def primenit_pravilo_obespecheniya_zayavki(nmcd: Decimal, ustanovleno_v_zayavke: bool, procent_v_zayavke: str) -> tuple[bool, list[str]]:
    """Правило 8 (раздел 7 ТЗ): при НМЦД до 5 000 000,00 руб. включительно
    обеспечение заявки не устанавливается, даже если в заявке указан процент."""
    preduprezhdeniya: list[str] = []
    if ustanovleno_v_zayavke and procent_v_zayavke and nmcd <= NMCD_GRANICA_OBESP_ZAYAVKI:
        preduprezhdeniya.append(
            f"в заявке указано обеспечение заявки {procent_v_zayavke}%, но НМЦД "
            f"{nmcd} руб. не выше 5 000 000,00 руб. — обеспечение заявки не устанавливается"
        )
        return False, preduprezhdeniya
    return ustanovleno_v_zayavke, preduprezhdeniya


def sverit_s_dopuskom(a: Decimal, b: Decimal, dopusk: Decimal = DOPUSK_NDS) -> bool:
    return abs(a - b) <= dopusk


def summa_poziciy(pozicii_kolichestvo_cena: list[tuple[Decimal, Decimal]]) -> Decimal:
    """НМЦД по позициям: сумма «количество x цена» (раздел 6 ТЗ)."""
    return okruglit(sum(okruglit(kolichestvo * cena) for kolichestvo, cena in pozicii_kolichestvo_cena))


def _naiti_stolbec(pozicii: list[dict], fragment: str) -> str | None:
    if not pozicii:
        return None
    for nazvanie in pozicii[0]:
        if fragment in nazvanie.lower():
            return nazvanie
    return None


def proverit_nmcd_po_poziciyam(pozicii: list[dict], nmcd: Decimal) -> list[str]:
    """Сверка НМЦД с суммой «количество x цена» по позициям (раздел 6 ТЗ)."""
    kol_vo_imya = _naiti_stolbec(pozicii, "кол-во")
    cena_imya = _naiti_stolbec(pozicii, "цена за ед")
    if not kol_vo_imya or not cena_imya:
        return []
    summa = summa_poziciy([
        (chislo(p[kol_vo_imya]["znachenie"]), chislo(p[cena_imya]["znachenie"])) for p in pozicii
    ])
    if not sverit_s_dopuskom(summa, nmcd, Decimal("0.01")):
        return [f"НМЦД {nmcd} не совпадает с суммой «количество x цена» по позициям ({summa})"]
    return []


def izvlech_stavku_nds(pozicii: list[dict]) -> Decimal | None:
    """Ставка НДС по позициям, если она одна и та же (столбец «НДС»);
    None — если не найдена, разная или товар не облагается НДС."""
    imya = _naiti_stolbec(pozicii, "ндс")
    if not imya:
        return None
    stavki = set()
    for p in pozicii:
        m = re.match(r"(\d+(?:[.,]\d+)?)\s*%", p[imya]["znachenie"])
        if not m:
            return None
        stavki.add(m.group(1).replace(",", "."))
    if len(stavki) != 1:
        return None
    return Decimal(stavki.pop())


def izvlech_summu_iz_nds_teksta(nds_tekst: str) -> Decimal | None:
    """Сумма НДС из фразы п. 8 («в т.ч. НДС 10% - 161 182,46 руб.»)."""
    m = re.search(r"(\d[\d\s ]*,\d{2})\s*руб", nds_tekst)
    return chislo(m.group(1)) if m else None


@dataclass
class OshibkaTarifnogoRascheta:
    stroka: dict
    soobshchenie: str


def proverit_tarifnyj_raschet(nmck_tarif: dict) -> list[str]:
    """Самопроверка арифметики тарифного метода (как в
    shablon/skripty_dlya_spravki/nmck_tarif.py): цена с надбавкой x количество
    = сумма строки; сумма строк = итог."""
    oshibki: list[str] = []
    pozicii = nmck_tarif.get("pozicii", [])
    itogo = nmck_tarif.get("itogo")
    nazvaniya = list(pozicii[0].keys()) if pozicii else []
    if len(nazvaniya) < 6:
        return oshibki
    kol_vo_imya, summa_imya, nadbavka_imya = nazvaniya[4], nazvaniya[5], nazvaniya[3]

    summa_stroк = Decimal(0)
    for poz in pozicii:
        kolichestvo = chislo(poz[kol_vo_imya])
        nadbavka = chislo(poz[nadbavka_imya])
        summa = chislo(poz[summa_imya])
        summa_stroк += summa
        dopusk = KOPEJKA * kolichestvo
        if abs(nadbavka * kolichestvo - summa) > dopusk:
            oshibki.append(
                f"позиция {poz.get(nazvaniya[0])}: {nadbavka} x {kolichestvo} != {summa} "
                f"(тарифный расчёт НМЦК)"
            )
    if itogo is not None and abs(summa_stroк - chislo(itogo)) > Decimal("0.02"):
        oshibki.append(f"сумма строк тарифного расчёта {summa_stroк} != итог {itogo}")
    return oshibki
