# -*- coding: utf-8 -*-
"""Извлечение данных карточки закупки из заявки в формате Word (раздел 9 ТЗ).

Те же правила, что для Excel (izvlechenie_excel.py): значения ищутся по
подписям, а не по номерам абзацев или таблиц — у разных подразделений заявки
устроены по-разному. Для каждого значения запоминается источник: абзац или
таблица/строка/столбец, и дословная цитата.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from docx import Document
from docx.text.paragraph import Paragraph

import pravila
from izvlechenie_excel import (
    Istochnik,
    Kartochka,
    format_chislo,
    naiti_zapis_podrazdeleniya,
    razobrat_cenu_vklyuchaet,
    s_nerazryvnymi_probelami_v_chislah,
)


@dataclass
class AbzacIstochnika:
    nomer: int
    tekst: str


def _abzacy(doc: Document) -> list[str]:
    return [p.text for p in doc.paragraphs]


def naiti_abzac_po_prefiksu(abzacy: list[str], prefiks: str) -> int | None:
    prefiks_low = prefiks.lower()
    for i, a in enumerate(abzacy):
        if a.strip().lower().startswith(prefiks_low):
            return i
    return None


def naiti_abzac_s_podpisyu(abzacy: list[str], *fragmenty: str) -> int | None:
    frag_low = [f.lower() for f in fragmenty]
    for i, a in enumerate(abzacy):
        nizhnij = a.lower()
        if nizhnij and all(f in nizhnij for f in frag_low):
            return i
    return None


def tekst_posle_podpisi(abzac: str, podpis: str) -> str:
    """Текст абзаца после подписи и двоеточия («Место доставки: ...»)."""
    i = abzac.lower().find(podpis.lower())
    if i == -1:
        return abzac.strip()
    ostatok = abzac[i + len(podpis):]
    return ostatok[1:].strip() if ostatok.startswith(":") else ostatok.strip()


def sobrat_abzacy_do(abzacy: list[str], posle: int, stop_prefiks: str, sohranyat_pustye: bool = False) -> list[str]:
    """Абзацы после данного индекса, пока очередной не начинается с
    указанного префикса (его не включать). Пустые абзацы — только если
    sohranyat_pustye=True (раздел 4 ТЗ: инструкция хранит пустые строки)."""
    rezultat = []
    for a in abzacy[posle + 1:]:
        if a.strip().lower().startswith(stop_prefiks.lower()):
            break
        if a.strip() or sohranyat_pustye:
            rezultat.append(a if a.strip() else "")
    return rezultat


class ZayavkaWord:
    """Заявка в файле Word (данные — в абзацах и таблицах)."""

    def __init__(self, put: str | Path):
        self.put = Path(put)
        self.fajl = self.put.name
        self.doc = Document(self.put)
        self.abzacy = _abzacy(self.doc)

    def _polozhit(self, karta: Kartochka, imya: str, mesto: str, znachenie: Any, citata: str = "") -> None:
        karta.znacheniya[imya] = znachenie
        karta.istochniki[imya] = Istochnik(self.fajl, mesto, "", citata)

    # ------------------------------------------------------------ предмет закупки
    def _predmet(self, karta: Kartochka) -> None:
        if not self.doc.tables:
            return
        t = self.doc.tables[0]
        for row in t.rows[1:]:
            tekst = row.cells[0].text.strip()
            prefiks = "Закрытый аукцион в электронной форме на "
            if not tekst.lower().startswith(prefiks.lower()):
                continue
            vin = tekst[len(prefiks):].strip()
            pervoe_slovo_im = {"поставку": "Поставка"}
            pervoe, _, ostatok = vin.partition(" ")
            im = pervoe_slovo_im.get(pervoe.lower(), pervoe.capitalize()) + " " + ostatok
            mesto = "таблица 1, строка 2, столбец 1"
            self._polozhit(karta, "predmet_vin", mesto, vin, tekst)
            self._polozhit(karta, "predmet_im", mesto, im, tekst)
            nmcd = row.cells[-1].text.strip()
            if nmcd:
                self._polozhit(karta, "nmcd", "таблица 1, строка 2, столбец 2", format_chislo(nmcd), nmcd)
            return

    # -------------------------------------------------------------------- НДС
    def _nds_tekst(self, karta: Kartochka) -> None:
        if len(self.doc.tables) < 2:
            return
        t = self.doc.tables[1]
        for row in t.rows:
            metka = row.cells[0].text.strip()
            m = re.match(r"В том числе НДС,?\s*(\d+(?:[.,]\d+)?)\s*%", metka, re.I)
            if not m:
                continue
            summa = row.cells[-1].text.strip()
            if not summa:
                continue
            tekst = f"в том числе НДС {m.group(1)}% - {format_chislo(summa)} руб."
            self._polozhit(karta, "nds_tekst", "таблица 2", tekst, f"{metka} | {summa}")
            # ставка НДС — для столбца «НДС» таблицы ТЗ, если его нет в заявке
            # (добавляется из условий заявки — раздел 5 ТЗ)
            self._polozhit(karta, "_stavka_nds_dlya_tablicy", "таблица 2", f"{m.group(1)}%", metka)
            return

    # ------------------------------- гарантийный срок (для столбца таблицы ТЗ)
    def _garantijnyj_srok(self, karta: Kartochka) -> None:
        i = naiti_abzac_s_podpisyu(self.abzacy, "гарантийный срок", "поставщика")
        if i is None:
            i = naiti_abzac_po_prefiksu(self.abzacy, "Гарантийный срок")
        if i is None:
            return
        m = re.search(r"не менее\s+\d+\s*месяц\w*", self.abzacy[i], re.I)
        if m:
            self._polozhit(karta, "_garantijnyj_srok_dlya_tablicy", f"абзац {i}", m.group(), self.abzacy[i])

    # ------------------------------------------------------- площадка (подсказка)
    def _ploshchadka_podskazka(self, karta: Kartochka) -> None:
        i = naiti_abzac_po_prefiksu(self.abzacy, "Электронная площадка:")
        if i is None:
            return
        tekst = tekst_posle_podpisi(self.abzacy[i], "Электронная площадка")
        self._polozhit(karta, "_ploshchadka_podskazka", f"абзац {i}", tekst, self.abzacy[i])

    # --------------------------------------------------- номер и дата заявки
    def _rekvizity_zayavki(self, karta: Kartochka) -> None:
        i = naiti_abzac_s_podpisyu(self.abzacy, "заявка", "№")
        if i is None:
            return
        tekst = self.abzacy[i].strip()
        j = naiti_abzac_po_prefiksu(self.abzacy, "Дата составления:")
        if j is not None:
            tekst = f"{tekst}; {self.abzacy[j].strip()}"
        self._polozhit(karta, "_nomer_i_data_zayavki", f"абзац {i}", tekst)

    # ---------------------------------------- простые подписи «label: значение»
    def _prostye_podpisi(self, karta: Kartochka) -> None:
        for podpis, imya in (
            ("Место доставки:", "mesto_postavki"),
            ("Сроки поставки:", "sroki_postavki"),
            ("Сроки и условия оплаты:", "usloviya_oplaty"),
        ):
            i = naiti_abzac_po_prefiksu(self.abzacy, podpis)
            if i is None:
                continue
            tekst = tekst_posle_podpisi(self.abzacy[i], podpis.rstrip(":"))
            self._polozhit(karta, imya, f"абзац {i}", tekst, self.abzacy[i])

        i = naiti_abzac_po_prefiksu(self.abzacy, "Срок действия договора:")
        if i is not None:
            tekst = tekst_posle_podpisi(self.abzacy[i], "Срок действия договора")
            self._polozhit(karta, "srok_dejstviya_dogovora", f"абзац {i}", tekst, self.abzacy[i])

    # ---------------------------------------------------------- цена включает
    def _cena_vklyuchaet(self, karta: Kartochka) -> None:
        i = naiti_abzac_po_prefiksu(self.abzacy, "Цена товара включает:")
        if i is None:
            return
        tekst = tekst_posle_podpisi(self.abzacy[i], "Цена товара включает")
        self._polozhit(karta, "cena_vklyuchaet", f"абзац {i}", razobrat_cenu_vklyuchaet(tekst), self.abzacy[i])

    # ------------------------------------------------------------ обеспечение
    def _obespechenie(self, karta: Kartochka) -> None:
        i = naiti_abzac_po_prefiksu(self.abzacy, "Обеспечение заявки и договора:")
        if i is None:
            return
        tekst = self.abzacy[i].strip()
        self._polozhit(karta, "obespechenie_tekst", f"абзац {i}", tekst)

        zayavka_ne = re.search(r"обеспечение заявки\s*[-–]\s*(не требуется|не применяется)", tekst, re.I)
        self._polozhit(karta, "obesp_zayavki", f"абзац {i}", {"ustanovleno": not bool(zayavka_ne)}, tekst)

        dogovor_procent = re.search(r"обеспечение договора\s*[-–]\s*(\d+(?:[.,]\d+)?)\s*%", tekst, re.I)
        dogovor_ne = re.search(r"обеспечение договора\s*[-–]\s*(не требуется|не применяется)", tekst, re.I)
        if dogovor_procent:
            self._polozhit(karta, "obesp_isp_procent", f"абзац {i}", dogovor_procent.group(1).replace(".", ","), tekst)
        elif dogovor_ne:
            self._polozhit(karta, "obesp_isp_procent", f"абзац {i}", "", tekst)

    # ------------------------------------------------- ответственный за договор
    def _otv_dogovor(self, karta: Kartochka) -> None:
        i = naiti_abzac_s_podpisyu(self.abzacy, "ответственный за заключение контракта")
        if i is None:
            return
        tekst = tekst_posle_podpisi(self.abzacy[i], "Ответственный за заключение контракта").strip().rstrip(".")
        self._polozhit(karta, "otv_dogovor", f"абзац {i}", tekst, self.abzacy[i])
        tel = re.search(r"телефон:?\s*([+0-9()\-\s]{7,})", tekst)
        if tel:
            self._polozhit(karta, "otv_dogovor_telefon", f"абзац {i}", tel.group(1).strip().rstrip(","), tekst)

    # --------------------------------------------------- требования к участникам
    def _trebovaniya_k_uchastnikam(self, karta: Kartochka, standartnye: set[str]) -> None:
        i = naiti_abzac_po_prefiksu(self.abzacy, "Требования к участникам")
        if i is None:
            return
        stroki = sobrat_abzacy_do(self.abzacy, i, "Начальная максимальная цена")
        self._polozhit(karta, "trebovaniya_k_uchastnikam", f"абзацы {i + 1}..", stroki, "\n".join(stroki))

        dopolnitelnye = [s for s in stroki if not pravila.sovpadaet_so_standartnym(s, standartnye)]
        if dopolnitelnye:
            tekst = dopolnitelnye[0].rstrip(".;")
            if not tekst.endswith(";"):
                tekst += ";"
            self._polozhit(karta, "dok_sootv_uchastnika", f"абзацы {i + 1}..", "требуется")
            self._polozhit(karta, "dok_sootv_tekst", f"абзацы {i + 1}..", tekst, dopolnitelnye[0])
        else:
            self._polozhit(karta, "dok_sootv_uchastnika", f"абзацы {i + 1}..", "не требуется")
            self._polozhit(karta, "dok_sootv_tekst", f"абзацы {i + 1}..", "")

    # ------------------------------------------------------------- инструкция
    def _instrukciya(self, karta: Kartochka) -> None:
        i = naiti_abzac_po_prefiksu(self.abzacy, "ИНСТРУКЦИЯ ПО ОПИСАНИЮ ТОВАРА")
        if i is None:
            return
        abzacy = sobrat_abzacy_do(self.abzacy, i, "Приложение к заявке № 2", sohranyat_pustye=True)
        self._polozhit(karta, "instrukciya_tovara", f"абзацы {i + 1}..", abzacy, "\n".join(abzacy))

    # ------------------------------------------------- утверждающий, согласующие
    def _utverzhdenie(self, karta: Kartochka, podpisanty: dict[str, tuple[str, str]]) -> None:
        zapis = podpisanty.get("utverzhdayuschij_po_umolchaniyu")
        if zapis is None:
            return
        dolzhnost, fio = zapis
        karta.znacheniya["utverzhdayu"] = {"dolzhnost": dolzhnost, "fio": fio}
        karta.istochniki["utverzhdayu"] = Istochnik("spravochnik", "podpisanty", "", "справочник подписантов, по умолчанию")

    def _soglasovanie(self, karta: Kartochka, podrazdeleniya: dict[str, tuple[str, str]]) -> None:
        # лист согласования заявки (таблица «СОГЛАСОВАНО») — только для
        # проверки, что выбранный согласующий там есть (самопроверка)
        tablica = self._tablica_soglasovano()
        if tablica is not None:
            karta.znacheniya["_list_soglasovaniya"] = list(dict.fromkeys(
                c.text.strip() for row in tablica.rows for c in row.cells if c.text.strip()))
            karta.istochniki["_list_soglasovaniya"] = Istochnik(self.fajl, "", "", "СОГЛАСОВАНО")
        i = naiti_abzac_po_prefiksu(self.abzacy, "Наименование подразделения:")
        tekst_podrazdeleniya = tekst_posle_podpisi(self.abzacy[i], "Наименование подразделения") if i is not None else None
        zapis = naiti_zapis_podrazdeleniya(tekst_podrazdeleniya, podrazdeleniya) if tekst_podrazdeleniya else None
        if zapis is None:
            return
        dolzhnost, fio = zapis
        karta.znacheniya["soglasovano"] = [{"dolzhnost": dolzhnost, "fio": fio}]
        karta.istochniki["soglasovano"] = Istochnik("spravochnik", "podrazdelenie_soglasuyushchij", "", f"подразделение: {tekst_podrazdeleniya}")


    def _tablica_soglasovano(self):
        i = naiti_abzac_po_prefiksu(self.abzacy, "СОГЛАСОВАНО")
        if i is None:
            return None
        # ищем первую таблицу документа, идущую после этого абзаца
        for t in self.doc.tables:
            if t.rows and t.rows[0].cells[0].text.strip() and "Директор" in t.rows[0].cells[0].text:
                return t
        return None

    # ------------------------------------------------------------- расчёт НМЦК
    def _nmck_raschet(self, karta: Kartochka) -> None:
        if len(self.doc.tables) < 3:
            return
        t = self.doc.tables[2]
        obekt = metod = ""
        kp: list[dict[str, str]] = []
        raschet: dict[str, Any] = {"smeta": False, "po_minimalnomu": False}
        obrazec = re.compile(r"(\d+)\s+от\s+(\d{2}\.\d{2}\.\d{4})")
        for row in t.rows:
            metka = row.cells[0].text.strip()
            if metka.startswith("Основные характеристики"):
                obekt = row.cells[1].text.strip()
            elif metka.startswith("Используемый метод"):
                metod = row.cells[1].text.strip()
            m = obrazec.search(row.cells[-2].text.strip()) if len(row.cells) >= 4 else None
            if m:
                kp.append({"vh_nomer": m.group(1), "vh_data": f"{m.group(2)}г.", "summa": format_chislo(row.cells[-1].text.strip())})
            poslednyaya = row.cells[-2].text.strip().lower()
            if poslednyaya.startswith("средняя арифметическая"):
                raschet["srednyaya"] = format_chislo(row.cells[-1].text.strip())
            elif poslednyaya.startswith("кол-во значений"):
                raschet["chislo_znachenij"] = format_chislo(row.cells[-1].text.strip())
            elif poslednyaya.startswith("кол-во (объем)"):
                raschet["obem"] = format_chislo(row.cells[-1].text.strip())
            elif poslednyaya.startswith("σ"):
                raschet["sko"] = format_chislo(row.cells[-1].text.strip())
            elif poslednyaya.startswith("коэффициент вариации"):
                raschet["koef_var"] = format_chislo(row.cells[-1].text.strip())
            elif poslednyaya.startswith("совокупность значений"):
                raschet["odnorodnost"] = row.cells[-1].text.strip().upper()

        raschet["po_obshchej_summe"] = bool(kp)
        if kp:
            raschet["obekt"] = obekt
            raschet["kp"] = kp
        self._polozhit(karta, "nmck_raschet", "таблица 3", raschet, metod)

    # ------------------------------------------------------------------ позиции
    def _pozicii(self, karta: Kartochka) -> None:
        if len(self.doc.tables) < 2:
            return
        t = self.doc.tables[1]
        zagolovki = [c.text.strip() for c in t.rows[0].cells]
        for row in t.rows[1:]:
            nomer = row.cells[0].text.strip()
            if not nomer.isdigit():
                continue
            poziciya: dict[str, Any] = {}
            for col, nazvanie in enumerate(zagolovki):
                if nazvanie in poziciya:
                    continue  # повторяющееся название объединённого столбца
                poziciya[nazvanie] = {
                    "znachenie": row.cells[col].text.strip(),
                    "fajl": self.fajl,
                    "list": "таблица 2",
                    "yacheyka": f"строка {nomer}, столбец «{nazvanie}»",
                }
            karta.pozicii.append(poziciya)

    # ----------------------------------------------- нацрежим по позициям (п. 5, 17)
    def _nacrezhim_po_poziciyam(self, karta: Kartochka) -> None:
        stolbec = next(
            (n for n in karta.pozicii[0] if "1875" in n), None,
        ) if karta.pozicii else None
        if stolbec is None:
            return
        mery = []
        for p in karta.pozicii:
            nomer = int(p["№"]["znachenie"])
            mera, poz = pravila.razobrat_stolbec_pp1875(p[stolbec]["znachenie"])
            mery.append((nomer, mera, poz))
        rezultat = pravila.rasschitat_nacrezhim_po_poziciyam(mery)
        for imya, znachenie in rezultat.items():
            karta.znacheniya[imya] = znachenie
            karta.istochniki[imya] = Istochnik(self.fajl, "таблица 2", f"столбец «{stolbec}»", "рассчитано по мерам позиций")

    def izvlechi(
        self,
        podpisanty: dict[str, tuple[str, str]] | None = None,
        podrazdeleniya: dict[str, tuple[str, str]] | None = None,
        trebovaniya_standartnye: set[str] | None = None,
    ) -> Kartochka:
        karta = Kartochka()
        self._predmet(karta)
        self._nds_tekst(karta)
        self._garantijnyj_srok(karta)
        self._ploshchadka_podskazka(karta)
        self._rekvizity_zayavki(karta)
        self._prostye_podpisi(karta)
        self._cena_vklyuchaet(karta)
        self._utverzhdenie(karta, podpisanty or {})
        self._soglasovanie(karta, podrazdeleniya or {})
        self._obespechenie(karta)
        self._otv_dogovor(karta)
        self._trebovaniya_k_uchastnikam(karta, trebovaniya_standartnye or set())
        self._instrukciya(karta)
        self._nmck_raschet(karta)
        self._pozicii(karta)
        self._nacrezhim_po_poziciyam(karta)
        return karta


def izvlechi_kartochku(
    put_zayavki: str | Path,
    put_podpisantov: str | Path | None = None,
    put_podrazdelenij: str | Path | None = None,
    put_trebovanij_standartnyh: str | Path | None = None,
) -> Kartochka:
    """Извлечь карточку закупки из заявки Word (раздел 9 ТЗ)."""
    from izvlechenie_excel import zagruzit_podpisantov, zagruzit_podrazdeleniya, zagruzit_trebovaniya_standartnye

    podpisanty = zagruzit_podpisantov(put_podpisantov) if put_podpisantov else {}
    podrazdeleniya = zagruzit_podrazdeleniya(put_podrazdelenij) if put_podrazdelenij else {}
    standartnye = zagruzit_trebovaniya_standartnye(put_trebovanij_standartnyh) if put_trebovanij_standartnyh else set()
    return ZayavkaWord(put_zayavki).izvlechi(podpisanty, podrazdeleniya, standartnye)
