# -*- coding: utf-8 -*-
"""Ядро модуля: шесть функций, которые вызывает программа (раздел 14 ТЗ).

Ядро не знает ни про окно, ни про браузер, ни про сеть — это даёт
lokalnyj_vhod/server.py. Карточка закупки хранится как JSON в
dannye/zakupki/<id>/kartochka.json — поэтому после перезапуска программы
незавершённая закупка открывается заново.

Шесть функций:
  opisanie_polej        — список полей карточки для текущей процедуры
  otkryt_kartochku       — создать карточку в папке закупки или открыть сохранённую
  zapolnit_iz_zayavki    — заполнить пустые поля из документов заявки
  izmenit_pole           — изменить одно поле и пересчитать зависимые
  proverit                — самопроверка: ошибки и предупреждения
  sformirovat             — собрать Word в папку закупки
"""
from __future__ import annotations

import calendar as calendar_std
import json
import logging
import re
import uuid
from dataclasses import asdict
from functools import lru_cache
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import openpyxl

import dopolnenie_pravilami
import rol_fajla
import sborka
import staryj_format
import sverka_otdelnyh_fajlov as sverka
from daty import (
    DatyIzveshcheniya,
    DataRazmeshcheniyaNerabochaya,
    NetKalendaryaNaGod,
    format_dannye_dlya_shablona,
    rasschitat_daty,
    rasschitat_ostalnye_daty,
    zagruzit_kalendar,
)
from izvlechenie_excel import Istochnik, format_chislo, izvlechi_kartochku as izvlechi_excel
from izvlechenie_word import izvlechi_kartochku as izvlechi_word
from samoproverka import proverit_kartochku, zagruzit_slovar
import tablica_tz
from tablica_tz import dobavit_otsutstvuyushchie_stolbcy, sobrat_tablicu_tz

MESYACY_RODITELNYJ = ["", "января", "февраля", "марта", "апреля", "мая", "июня",
                       "июля", "августа", "сентября", "октября", "ноября", "декабря"]
MESYACY_IMENITELNYJ = ["", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

ROLI_PODPISI = {
    "zayavka": "Заявка",
    "tz": "Техническое задание",
    "raschet_nmck": "Расчёт НМЦК",
    "instrukciya": "Инструкция по описанию товара",
    "smeta": "Смета",
    "ne_ispolzuetsya": "Не используется",
}

SPRAVOCHNIK_FORM_DOGOVORA = "Формы договора.xlsx"

# где искать возможные опечатки в тексте заявки (решение специалиста от
# 11.10.2026): предмет, текст о нацрежиме, наименования позиций ТЗ
_PODPISI_POLEJ_OPECHATOK = {
    "predmet_im": "предмет закупки (п. 6)",
    "predmet_vin": "предмет на титульном листе",
    "nacrezhim_tekst": "текст о нацрежиме (п. 5)",
    "instrukciya_tovara": "инструкция по описанию товара",
}


def _slovar_opechatok(nastrojki: "Nastrojki") -> set[str]:
    put = nastrojki.sp("Словарь для проверки опечаток.xlsx")
    return zagruzit_slovar(put) if put.exists() else set()


def _mestah(n: int) -> str:
    if n % 10 == 1 and n % 100 != 11:
        return "месте"
    return "местах"


def _stolbec_naimenovaniya(poziciya: dict) -> str | None:
    return next((k for k in poziciya if "наименование" in k.lower()), None)


def opechatki_zayavki(nastrojki: "Nastrojki", sostoyanie: dict) -> list[dict]:
    """Возможные опечатки в тексте заявки, по словам: {najdeno, ispravlenie,
    mesta: [где встречается], gruppy: {"predmet", "nacrezhim", "pozicii"}}.
    Слова, которые сотрудник оставил «как в заявке», не показываются."""
    from samoproverka import naiti_vozmozhnye_opechatki

    karta = sostoyanie.get("karta") or {}
    znach = karta.get("znacheniya") or {}
    teksty: list[tuple[str, str, str]] = []  # (текст, где, группа)
    if znach.get("predmet_im"):
        teksty.append((str(znach["predmet_im"]), "предмет закупки (п. 6)", "predmet"))
    if znach.get("nacrezhim_tekst"):
        teksty.append((str(znach["nacrezhim_tekst"]), "текст о нацрежиме (п. 5)", "nacrezhim"))
    for i, p in enumerate(karta.get("pozicii") or [], start=1):
        stolbec = _stolbec_naimenovaniya(p)
        if stolbec and (p[stolbec] or {}).get("znachenie"):
            teksty.append((str(p[stolbec]["znachenie"]), f"позиция {i} ТЗ", "pozicii"))

    ostavleny = {z.split(":", 1)[-1].lower() for z in (sostoyanie.get("opechatki_ostavleny") or [])}
    slovar = _slovar_opechatok(nastrojki)
    po_slovam: dict[str, dict] = {}
    for tekst, gde, gruppa in teksty:
        for o in naiti_vozmozhnye_opechatki([tekst], slovar):
            klyuch = o.najdeno.lower()
            if klyuch in ostavleny:
                continue
            zapis = po_slovam.setdefault(klyuch, {"najdeno": o.najdeno, "ispravlenie": o.ispravlenie,
                                                  "mesta": [], "gruppy": []})
            if gde not in zapis["mesta"]:
                zapis["mesta"].append(gde)
            if gruppa not in zapis["gruppy"]:
                zapis["gruppy"].append(gruppa)
    return list(po_slovam.values())


def _ispravit_slovo_vezde(karta: dict, najdeno: str, ispravlenie: str) -> list[str]:
    """Заменяет слово во всех полях из заявки и в позициях ТЗ; источник —
    «исправлено сотрудником, в заявке: <как было>». Возвращает, где исправлено."""
    obrazec = re.compile(rf"(?<!\w){re.escape(najdeno)}(?!\w)")
    zamena = lambda _: ispravlenie  # noqa: E731 — подставить буквально, без разбора «\…»
    mesta: list[str] = []

    def istochnik(bylo) -> dict:
        bylo_tekst = "\n".join(bylo) if isinstance(bylo, list) else str(bylo)
        return {"fajl": "sotrudnik", "list_": "", "yacheyka": "",
                "citata": f"исправлено сотрудником, в заявке: {bylo_tekst}"}

    for imya, bylo in list(karta["znacheniya"].items()):
        if imya.startswith("_"):
            continue
        if isinstance(bylo, str) and obrazec.search(bylo):
            karta["znacheniya"][imya] = obrazec.sub(zamena, bylo)
        elif isinstance(bylo, list) and any(isinstance(s, str) and obrazec.search(s) for s in bylo):
            karta["znacheniya"][imya] = [obrazec.sub(zamena, s) if isinstance(s, str) else s for s in bylo]
        else:
            continue
        karta.setdefault("istochniki", {})[imya] = istochnik(bylo)
        mesta.append(_PODPISI_POLEJ_OPECHATOK.get(imya) or _podpis_polya(imya) or imya)
    for i, p in enumerate(karta.get("pozicii") or [], start=1):
        izmeneno = False
        for yach in p.values():
            tekst = yach.get("znachenie") if isinstance(yach, dict) else None
            if isinstance(tekst, str) and obrazec.search(tekst):
                yach["zayavka_bylo"] = tekst
                yach["znachenie"] = obrazec.sub(zamena, tekst)
                yach["fajl"] = "sotrudnik"
                yach["citata"] = f"исправлено сотрудником, в заявке: {tekst}"
                izmeneno = True
        if izmeneno:
            mesta.append(f"позиция {i} ТЗ")
    # расчёт НМЦК по позициям (Приложение № 2)
    nmck_poz = karta["znacheniya"].get("_nmck_poz")
    if nmck_poz and any(obrazec.search(s) for stroka in nmck_poz["stroki"] for s in stroka):
        nmck_poz["stroki"] = [[obrazec.sub(zamena, s) for s in stroka] for stroka in nmck_poz["stroki"]]
        mesta.append("расчёт НМЦК (приложение № 2)")
    # таблицы отдельного файла ТЗ исправляются при сборке документа
    karta.setdefault("ispravleniya_opechatok", {})[najdeno] = ispravlenie
    return mesta


# поля извещения, без которых документ неполный: (поле, подпись, раздел карточки)
OBYAZATELNYE_IZ_ZAYAVKI = (
    ("predmet_im", "предмет закупки", "s2"),
    ("nmcd", "НМЦК", "s3"),
    ("nds_tekst", "НДС (п. 8)", "s3"),
    ("otv_dogovor", "ответственный за исполнение договора", "s5"),
    ("instrukciya_tovara", "инструкция по описанию товара (приложение № 4)", "s8"),
)

ETAPY_ZAKUPKI = {
    "chernovik": "Черновик",
    "na_soglasovanii": "На согласовании",
    "razmeshchenie": "Размещение",
}

CHTO_BERYOTSYA_PO_ROLI = {
    "zayavka": "Предмет, позиции, цены, НДС, условия, обеспечение, нацрежим, ответственный, инструкция (прил. № 1), расчёт НМЦК (прил. № 2)",
    "tz": "Позиции технического задания — сверяются с заявкой",
    "raschet_nmck": "Расчёт НМЦК (Приложение № 2) — сверяется с заявкой",
    "instrukciya": "Инструкция по описанию товара (Приложение № 1) — сверяется с заявкой",
    "smeta": "Смета — в извещение не добавляется",
    "ne_ispolzuetsya": "не читается программой",
}


def _klyuch_alfavita(fio: str) -> str:
    """Для сортировки по алфавиту: «ё» — как «е» (решение специалиста
    от 10.10.2026)."""
    return fio.replace("Ё", "Е").replace("ё", "е")


def sotrudniki_po_alfavitu(sotrudniki: list[str] | dict[str, str], zapomnennyj: str | None = None) -> list[str]:
    """Сотрудники по алфавиту (фамилия, имя, отчество — одной строкой,
    «ё» = «е»); запомненный на этом компьютере — первым (раздел 4 ТЗ)."""
    vse = list(sotrudniki)
    ostalnye = sorted((f for f in vse if f != zapomnennyj), key=_klyuch_alfavita)
    if zapomnennyj and zapomnennyj in vse:
        return [zapomnennyj] + ostalnye
    return ostalnye


def sokratit_fio(fio: str) -> str:
    """«Миранова Валентина Станиславовна» -> «Миранова В.С.» (шапка страницы,
    как на макете)."""
    chasti = fio.split()
    if len(chasti) < 2:
        return fio
    familiya, *imena = chasti
    return familiya + " " + "".join(f"{i[0]}." for i in imena if i)


def nomer_iz_nazvaniya_papki(imya_papki: str) -> str:
    """Номер закупки — первое слово названия папки, если оно начинается с
    цифры («512зэа КЦ 1961-а-26» -> «512зэа»); иначе — пусто, сотрудник
    впишет вручную (раздел 4 ТЗ: «нет номера в названии папки — впишите вручную»)."""
    pervoe_slovo = imya_papki.strip().split(" ", 1)[0] if imya_papki.strip() else ""
    return pervoe_slovo if pervoe_slovo and pervoe_slovo[0].isdigit() else ""


def format_nomer_i_data_zayavki(tekst: str | None) -> str | None:
    """«№ 1961/а-26 от «15» июля 2026 г.» -> «№ 1961/а-26 от 15.07.2026» —
    короткая дата для шапки карточки, таблицы ТЗ и списка закупок."""
    if not tekst:
        return tekst
    mesyacy = {m: i for i, m in enumerate(MESYACY_RODITELNYJ) if m}

    def _zamena(m: re.Match) -> str:
        den, mesyac_imya, god = m.group(1), m.group(2).lower(), m.group(3)
        nomer_mesyaca = mesyacy.get(mesyac_imya)
        if not nomer_mesyaca:
            return m.group(0)
        return f"от {int(den):02d}.{nomer_mesyaca:02d}.{god}"

    return re.sub(r'от\s+["«]?(\d{1,2})["»]?\s+(\S+)\s+(\d{4})\s*г?\.?', _zamena, tekst)


def zagruzit_ploshchadki(nastrojki: Nastrojki) -> dict[str, str]:
    put = nastrojki.sp("Электронные площадки.xlsx")
    if not put.exists():
        return {}
    wb = openpyxl.load_workbook(put, data_only=True)
    ws = wb["Площадки"]
    rezultat = {}
    for adres, nazvanie, _ in ws.iter_rows(min_row=2, max_col=3, values_only=True):
        if adres:
            rezultat[str(adres).strip()] = f"{nazvanie} — {adres}"
    return rezultat


def dobavit_fajly_v_papku(fajly: list[str], put_papki: str | Path | None = None) -> str | None:
    """«Добавить файлы» (раздел 9 ТЗ): копирует выбранные файлы в папку, где
    лежит файл с ролью «Заявка» (или в put_papki, если он уже известен).
    Возвращает путь к папке, куда сложили файлы, или None, если файлов нет."""
    import shutil
    fajly_put = [Path(f) for f in fajly if f]
    if not fajly_put:
        return str(put_papki) if put_papki else None

    if put_papki is None:
        zayavka = next((f for f in fajly_put if rol_fajla.opredelit_rol(f) == "zayavka"), None)
        put_papki = (zayavka or fajly_put[0]).parent
    put_papki = Path(put_papki)

    for f in fajly_put:
        if f.parent == put_papki:
            continue
        naznachenie = put_papki / f.name
        if naznachenie.exists():
            continue
        try:
            shutil.copy2(f, naznachenie)
        except OSError:
            pass
    return str(put_papki)


def pereschitat_dokumenty(nastrojki: Nastrojki, id_zakupki: str) -> dict:
    """Пересобирает список файлов закупки после «Добавить файлы» (раздел 9 ТЗ)."""
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    if sostoyanie is None:
        raise ValueError(f"закупка {id_zakupki} не найдена")
    put_papki = Path(sostoyanie["put_papki"])
    sostoyanie["dokumenty"] = _sobrat_documenty(put_papki, nastrojki)
    _sohranit(nastrojki, id_zakupki, sostoyanie)
    return sostoyanie


def zagruzit_podpisantov_spiska(nastrojki: Nastrojki, rol: str | None = None) -> list[dict]:
    """Список подписантов титульного листа (справочник «Подписанты.xlsx») —
    должность + И.О. Фамилия. rol="utverzhdaet"/"soglasuet" — только эта роль
    (решение специалиста: столбец «Роль» в справочнике); без rol — все."""
    put = nastrojki.sp("Подписанты.xlsx")
    if not put.exists():
        return []
    from izvlechenie_excel import zagruzit_podpisantov
    spravochnik = zagruzit_podpisantov(put)
    if rol is None:
        vse = spravochnik.get("vse", [])
        return [{"dolzhnost": d, "fio": f} for d, f in vse]
    soglasovanie_rus = {"utverzhdaet": "утверждает", "soglasuet": "согласует"}.get(rol, rol)
    return [
        {"dolzhnost": p["dolzhnost"], "fio": p["fio"]}
        for p in spravochnik.get("vse_s_rolyu", []) if p["rol"] == soglasovanie_rus
    ]

def zagruzit_soglasuyushchih(nastrojki: Nastrojki) -> list[dict]:
    """Кого можно выбрать в СОГЛАСОВАНО: подписанты с ролью «согласует» и
    согласующие подразделений («Подразделение — согласующий.xlsx»), без повторов."""
    spisok = zagruzit_podpisantov_spiska(nastrojki, rol="soglasuet")
    put = nastrojki.sp("Подразделение — согласующий.xlsx")
    if put.exists():
        from izvlechenie_excel import zagruzit_podrazdeleniya
        for dolzhnost, fio in zagruzit_podrazdeleniya(put).values():
            if fio in ("", "None"):
                continue
            zapis = {"dolzhnost": dolzhnost, "fio": fio}
            if zapis not in spisok:
                spisok.append(zapis)
    return spisok

# ------------------------------------------------------------------ настройки

OZHIDANIE_SETEVOJ_PAPKI = 2.0  # секунд: недоступная сетевая папка не задерживает запуск


def papka_dostupna(put: Path, ozhidanie: float = OZHIDANIE_SETEVOJ_PAPKI) -> bool:
    """Есть ли папка — с ожиданием не дольше ozhidanie секунд (Windows ищет
    недоступный сетевой сервер десятки секунд)."""
    import threading

    rezultat: list[bool] = []

    def proverit_papku():
        try:
            rezultat.append(put.is_dir())
        except OSError:
            rezultat.append(False)

    potok = threading.Thread(target=proverit_papku, daemon=True)
    potok.start()
    potok.join(ozhidanie)
    return bool(rezultat and rezultat[0])


@lru_cache(maxsize=1)
def tochki_avtonomno():
    """Заготовки пяти точек подключения — podklyuchenie_avtonomno/tochki_podklyucheniya.py."""
    import importlib.util

    put = Path(__file__).resolve().parents[1] / "podklyuchenie_avtonomno" / "tochki_podklyucheniya.py"
    spec = importlib.util.spec_from_file_location("tochki_podklyucheniya", put)
    modul = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modul)
    return modul


class Nastrojki:
    """Настройки (точка подключения 1) и доступ к остальным четырём точкам
    (раздел 14 ТЗ). tochki — объект с функциями raspoznannyj_tekst,
    dannye_proverki_zayavki, sopostavlenie_stolbcov, vyzov_ii и константой
    VYKLYUCHENO; по умолчанию — автономные заготовки. Ядро обращается к
    внешнему только через self.tochki."""

    def __init__(
        self, put_dannyh: str | Path, put_zakupok: str | Path | None = None, tochki=None,
        put_obshchih_spravochnikov: str | Path | None = None,
        put_obshchih_zakupok: str | Path | None = None,
    ):
        self.put_dannyh = Path(put_dannyh)
        self.put_spravochnikov = self.put_dannyh / "spravochniki"
        # общая (сетевая) папка справочников; недоступна — локальные и предупреждение
        self.preduprezhdenie_spravochniki = None
        if put_obshchih_spravochnikov:
            obshchie = Path(put_obshchih_spravochnikov)
            if papka_dostupna(obshchie):
                self.put_spravochnikov = obshchie
            else:
                self.preduprezhdenie_spravochniki = (
                    f"Общие справочники недоступны ({obshchie}), использую локальные."
                )
        mestnye_zakupki = Path(put_zakupok) if put_zakupok else self.put_dannyh / "zakupki"
        self.put_zakupok = mestnye_zakupki
        # общая (сетевая) папка закупок — руководитель видит закупки всех
        # сотрудников; недоступна — карточки только на этом компьютере
        self.obshchaya_papka_zakupok = False
        self.preduprezhdenie_zakupki = None
        if put_obshchih_zakupok:
            obshchaya = Path(put_obshchih_zakupok)
            if papka_dostupna(obshchaya):
                self.put_zakupok = obshchaya
                self.obshchaya_papka_zakupok = True
            else:
                self.preduprezhdenie_zakupki = (
                    "Общая папка закупок недоступна, закупки сохраняются только на этом компьютере."
                )
        self.put_zakupok.mkdir(parents=True, exist_ok=True)
        # временные копии файлов старых форматов (.xls/.doc/.rtf → .xlsx/.docx) — на этом компьютере
        self.put_vremennyh = mestnye_zakupki.parent / "vremennoe"
        self.tochki = tochki or tochki_avtonomno()

    def vklyucheno(self, otvet) -> bool:
        return otvet != self.tochki.VYKLYUCHENO

    def vyzvat_tochku(self, imya: str, *argumenty) -> tuple[str, Any]:
        """Вызов точки подключения 2–5 без ошибок для сотрудника: (состояние,
        ответ). Состояния: VYKLYUCHENO_SOSTOYANIE — точка не подключена;
        NEDOSTUPEN_SOSTOYANIE — подключена, но сервер не ответил (исключение
        или ответ tochki.NEDOSTUPEN); RABOTAET — ответ получен."""
        try:
            otvet = getattr(self.tochki, imya)(*argumenty)
        except Exception as e:  # сервер программы недоступен — модуль работает дальше сам
            logging.getLogger("podgotovka_izveshcheniya").info("точка подключения %s: сервер недоступен (%s)", imya, e)
            return NEDOSTUPEN_SOSTOYANIE, None
        if otvet == self.tochki.VYKLYUCHENO:
            return VYKLYUCHENO_SOSTOYANIE, None
        if otvet == getattr(self.tochki, "NEDOSTUPEN", object()):
            return NEDOSTUPEN_SOSTOYANIE, None
        return RABOTAET, otvet

    def sp(self, imya_fajla: str) -> Path:
        return self.put_spravochnikov / imya_fajla


# ------------------------------------------------------------------ описание полей

def opisanie_polej(procedura: str = "zakrytyj_aukcion") -> list[dict[str, Any]]:
    """Список полей карточки: имя, группа, подпись (метка пункта), тип,
    обязательность — раздел 4 ТЗ. Строится по разделам карточки; по нему
    обшивка строит экран."""
    def pole(imya, grupa, podpis, tip="text", obyazatelno=False):
        return {"imya": imya, "grupa": grupa, "podpis": podpis, "tip": tip, "obyazatelno": obyazatelno}

    return [
        pole("nomer_zakupki", "obshchie", "имя файла", obyazatelno=True),
        pole("predmet_im", "obshchie", "п. 6"),
        pole("predmet_vin", "obshchie", "титул"),
        pole("god", "obshchie", "титул"),
        pole("utverzhdayu", "obshchie", "титул", tip="podpisant"),
        pole("soglasovano", "obshchie", "титул", tip="spisok_podpisantov"),
        pole("kontakt_fio", "obshchie", "п. 1"),
        pole("kontakt_tel", "obshchie", "п. 1"),
        pole("ploshchadka", "obshchie", "п. 3", tip="vybor", obyazatelno=True),
        pole("instrukciya_tovara", "obshchie", "Прил. 4", tip="spisok_abzacev"),
        pole("nmcd", "cena", "п. 8"),
        pole("nds_tekst", "cena", "п. 8"),
        pole("cena_vklyuchaet", "cena", "п. 8", tip="spisok_abzacev"),
        pole("obesp_zayavki", "cena", "п. 19", tip="pereklyuchatel"),
        pole("obesp_isp_procent", "cena", "п. 20"),
        pole("obesp_isp_summa", "cena", "п. 20"),
        pole("otv_dogovor", "cena", "п. 32"),
        pole("nr_primenyaetsya", "nacrezhim", "п. 5", tip="pereklyuchatel"),
        pole("nr_zapret", "nacrezhim", "п. 5", tip="pereklyuchatel"),
        pole("nr_ogranichenie", "nacrezhim", "п. 5", tip="pereklyuchatel"),
        pole("nr_preimushchestvo", "nacrezhim", "п. 5", tip="pereklyuchatel"),
        pole("ru_minzdrav", "nacrezhim", "п. 17", tip="pereklyuchatel"),
        pole("dok_sootv_uchastnika", "nacrezhim", "п. 17", tip="pereklyuchatel"),
        pole("dok_sootv_tekst", "nacrezhim", "п. 17"),
        pole("nmck_raschet", "nmck", "Прил. 2", tip="raschet"),
        pole("d_razmeshcheniya", "daty", "п. 10, 12"),
        pole("d_zaprosy_do", "daty", "п. 10"),
        pole("d_podacha_do", "daty", "п. 12"),
        pole("t_podacha_do", "daty", "п. 12"),
        pole("d_rassmotrenie", "daty", "п. 13"),
        pole("d_aukcion", "daty", "п. 22"),
        pole("t_aukcion", "daty", "п. 22"),
        pole("d_itogi", "daty", "п. 23"),
        pole("forma_dogovora", "dogovor", "Прил. 5", tip="vybor", obyazatelno=True),
    ]


# ------------------------------------------------------------------ путь закупки

def _put_zakupki(nastrojki: Nastrojki, id_zakupki: str) -> Path:
    return nastrojki.put_zakupok / id_zakupki


def _put_kartochki(nastrojki: Nastrojki, id_zakupki: str) -> Path:
    return _put_zakupki(nastrojki, id_zakupki) / "kartochka.json"


def _sohranit(nastrojki: Nastrojki, id_zakupki: str, sostoyanie: dict) -> None:
    put = _put_kartochki(nastrojki, id_zakupki)
    put.parent.mkdir(parents=True, exist_ok=True)
    sostoyanie["izmeneno"] = datetime.now().strftime("%H:%M")
    put.write_text(json.dumps(sostoyanie, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


def _zagruzit(nastrojki: Nastrojki, id_zakupki: str) -> dict | None:
    put = _put_kartochki(nastrojki, id_zakupki)
    if not put.exists():
        return None
    return json.loads(put.read_text(encoding="utf-8"))


def spisok_zakupok(nastrojki: Nastrojki) -> list[dict]:
    """Закупки в работе — для стартового экрана (раздел 14 ТЗ: видны только
    «Черновик» и «На согласовании»)."""
    rezultat = []
    for papka in sorted(nastrojki.put_zakupok.iterdir()) if nastrojki.put_zakupok.exists() else []:
        s = _zagruzit(nastrojki, papka.name)
        if s:
            rezultat.append(s)
    return rezultat


# ------------------------------------------------------------------ сбор файлов

_SLUZHEBNYE_DOK = {"otv_dogovor_telefon", "_ploshchadka_podskazka", "_nomer_i_data_zayavki",
                    "_garantijnyj_srok_dlya_tablicy", "_stavka_nds_dlya_tablicy"}


class ZayavkaNeProchitana(ValueError):
    """Заявка не найдена или не прочитана — карточку не создавать (текст для сотрудника)."""


class KartochkaBezDannyh(ValueError):
    """В карточке нет данных заявки — действие невозможно (текст для сотрудника)."""

    def __init__(self, id_zakupki: str = ""):
        super().__init__(
            "В этой карточке нет данных заявки: заявка не была прочитана. "
            "Откройте папку закупки заново и укажите файл заявки."
        )
        self.id_zakupki = id_zakupki


def _karta_ili_oshibka(sostoyanie: dict) -> dict:
    karta = sostoyanie.get("karta")
    if not karta or karta.get("znacheniya") is None:
        raise KartochkaBezDannyh(sostoyanie.get("id", ""))
    return karta


def _dlya_chteniya(put: str | Path, nastrojki: Nastrojki) -> Path:
    """Файл для чтения: .xls/.doc — временная копия .xlsx/.docx во временной
    папке программы (staryj_format), остальные — сам файл."""
    return staryj_format.novaya_kopiya(put, nastrojki.put_vremennyh)


def sobrat_documenty(put_papki: str | Path, nastrojki: Nastrojki, roli: dict[str, str] | None = None) -> list[dict]:
    """Файлы папки закупки с ролями (раздел 9 ТЗ); roli — выбор сотрудника
    {путь: роль}, он важнее определённой программой."""
    dokumenty = _sobrat_documenty(Path(put_papki), nastrojki)
    for d in dokumenty:
        if roli and d["put"] in roli:
            d["rol"] = roli[d["put"]]
            d["rol_vybral_sotrudnik"] = True
    return dokumenty


# состояния функций, которые даёт сервер основной программы (точки 2–5)
VYKLYUCHENO_SOSTOYANIE, NEDOSTUPEN_SOSTOYANIE, RABOTAET = "vyklyucheno", "nedostupen", "rabotaet"
PODPIS_OCR = "Включается при подключении к серверу распознавания"
PODPIS_II = "Включается при подключении к серверу ИИ"
SERVER_NEDOSTUPEN = "Сервер недоступен, попробуйте позже"
PROVERKA_SVYAZI_II = "Проверка связи: ответьте «да»."
_sostoyanie_ii_kesh: dict = {}


def sostoyanie_ii(nastrojki: Nastrojki, svezhest: float = 60.0) -> str:
    """Подключён ли ИИ (точка 5) и отвечает ли сервер — для подписи на экране;
    ответ запоминается на svezhest секунд, чтобы не спрашивать сервер на каждой странице."""
    import time
    klyuch = id(nastrojki.tochki)
    zapis = _sostoyanie_ii_kesh.get(klyuch)
    if zapis and time.time() - zapis[1] < svezhest:
        return zapis[0]
    sostoyanie, _ = nastrojki.vyzvat_tochku("vyzov_ii", PROVERKA_SVYAZI_II)
    _sostoyanie_ii_kesh[klyuch] = (sostoyanie, time.time())
    return sostoyanie


def _sobrat_documenty(put_papki: Path, nastrojki: Nastrojki) -> list[dict]:
    """Файлы папки закупки с ролью по содержимому (раздел 9 ТЗ)."""
    rezultat = []
    for put in rol_fajla.sobrat_fajly_zakupki(put_papki):
        oshibka_chteniya = None
        try:
            rol = rol_fajla.opredelit_rol(_dlya_chteniya(put, nastrojki), imya=put.name)
        except staryj_format.NeUdalosProchitat as e:
            oshibka_chteniya = str(e)
            rol = rol_fajla.rol_po_imeni(put.name) or "ne_ispolzuetsya"
        dokument = {
            "put": str(put),
            "imya": put.name,
            "gde": str(put.parent.relative_to(put_papki)) if put.parent != put_papki else "в самой папке",
            "rol": rol,
        }
        if oshibka_chteniya:
            dokument["oshibka_chteniya"] = oshibka_chteniya
        if put.suffix.lower() == ".pdf":
            # точка подключения 2: распознанный текст PDF (автономно — выключено;
            # сервер недоступен — без ошибки, строка файла серая с подписью)
            sostoyanie, tekst = nastrojki.vyzvat_tochku("raspoznannyj_tekst", put)
            dokument["ocr"] = sostoyanie
            if sostoyanie == RABOTAET:
                dokument["raspoznannyj_tekst"] = tekst
        rezultat.append(dokument)
    return rezultat


def _izvlech_iz_zayavki(put_zayavki: Path, nastrojki: Nastrojki, polzovatel_fio: str | None, dokumenty: list[dict] | None = None):
    karta = _izvlech_iz_fajla_zayavki(put_zayavki, nastrojki, polzovatel_fio)
    if dokumenty:
        _dopolnit_iz_otdelnyh_fajlov(karta, dokumenty, nastrojki)

    # точка подключения 3: данные проверки заявки из программы — первыми,
    # остальное — из файла (раздел 14 ТЗ, правило 4)
    sostoyanie, gotovye = nastrojki.vyzvat_tochku("dannye_proverki_zayavki", put_zayavki)
    if sostoyanie == RABOTAET:
        karta.znacheniya.update(gotovye)

    # точка подключения 4: сопоставление столбцов таблицы позиций
    zagolovki = list(karta.pozicii[0].keys()) if karta.pozicii else []
    sostoyanie, sopostavlenie = nastrojki.vyzvat_tochku("sopostavlenie_stolbcov", zagolovki)
    if sostoyanie == RABOTAET:
        karta.pozicii = [{sopostavlenie.get(k, k): v for k, v in poz.items()} for poz in karta.pozicii]
    return karta


def _izvlech_iz_fajla_zayavki(put_zayavki: Path, nastrojki: Nastrojki, polzovatel_fio: str | None):
    put_zayavki = Path(put_zayavki)
    fajl = _dlya_chteniya(put_zayavki, nastrojki)
    if fajl.suffix.lower() == ".docx":
        karta = izvlechi_word(
            fajl, nastrojki.sp("Подписанты.xlsx"), nastrojki.sp("Подразделение — согласующий.xlsx"),
            nastrojki.sp("Требования к участникам (стандартные).xlsx"),
        )
    else:
        karta = izvlechi_excel(
            fajl, nastrojki.sp("Подписанты.xlsx"), nastrojki.sp("Подразделение — согласующий.xlsx"),
            nastrojki.sp("Требования к участникам (стандартные).xlsx"),
            nastrojki.sp("Сотрудники отдела.xlsx"), polzovatel_fio,
        )
    if fajl != put_zayavki:
        # источник — исходный файл сотрудника, а не временная копия
        for istochnik in karta.istochniki.values():
            if istochnik.fajl == fajl.name:
                istochnik.fajl = put_zayavki.name
    dopolnenie_pravilami.dopolnit(karta)
    return karta


def _dopolnit_iz_otdelnyh_fajlov(karta, dokumenty: list[dict], nastrojki: Nastrojki) -> None:
    """Чего нет в заявке, но есть в отдельном файле: инструкция по описанию
    товара — из файла инструкции или ТЗ (пример 426зэа: инструкция в файле ТЗ)."""
    if karta.znacheniya.get("instrukciya_tovara"):
        return
    for rol in ("instrukciya", "tz"):
        for d in dokumenty:
            if d["rol"] != rol or d.get("oshibka_chteniya"):
                continue
            try:
                put = _dlya_chteniya(d["put"], nastrojki)
                if put.suffix.lower() != ".docx":
                    continue
                abzacy = sverka.procitat_instrukciyu_iz_fajla(put, tolko_s_zagolovkom=True)
            except Exception:
                continue
            # дословно, с пустыми абзацами между пунктами (пробелы по краям абзаца —
            # не текст); пустые в начале и в конце — не нужны
            abzacy = [a.strip() for a in abzacy]
            while abzacy and not abzacy[-1].strip():
                abzacy.pop()
            while abzacy and not abzacy[0].strip():
                abzacy.pop(0)
            if abzacy:
                karta.znacheniya["instrukciya_tovara"] = abzacy
                karta.istochniki["instrukciya_tovara"] = Istochnik(d["imya"], "", "", "\n".join(abzacy)[:300])
                return


def _prochitat_zayavku(dokument: dict, nastrojki: Nastrojki, polzovatel_fio: str | None, dokumenty: list[dict] | None = None):
    """Читает заявку; не получилось — ZayavkaNeProchitana с причиной."""
    imya = dokument["imya"]
    if dokument.get("oshibka_chteniya"):
        raise ZayavkaNeProchitana(dokument["oshibka_chteniya"])
    try:
        karta = _izvlech_iz_zayavki(Path(dokument["put"]), nastrojki, polzovatel_fio, dokumenty)
    except staryj_format.NeUdalosProchitat as e:
        raise ZayavkaNeProchitana(str(e))
    except Exception:
        # файл не открылся (испорчен) — сотруднику без технических подробностей
        raise ZayavkaNeProchitana(str(staryj_format.NeUdalosProchitat(imya)))
    if not karta.znacheniya.get("predmet_im") and not karta.pozicii:
        raise ZayavkaNeProchitana(
            f"Заявку прочитать не удалось: {imya} — в файле не найдены предмет закупки и позиции. "
            "Проверьте, что это заявка, или укажите другой файл."
        )
    return karta


ROLI_CHITAEMYH = {".xlsx", ".xls", ".docx", ".doc", ".rtf"}


def proverit_papku(
    nastrojki: Nastrojki, put_papki: str | Path, polzovatel_fio: str | None = None, roli: dict[str, str] | None = None,
) -> dict:
    """Что решить перед созданием карточки (экран «Новая закупка», раздел 9 ТЗ):
    oshibka_zayavki — «Заявка не найдена» / «Заявку прочитать не удалось: …»;
    neopredeleny — файлы Excel/Word, роль которых не определена;
    rashozhdeniya — отдельные файлы не совпали с заявкой;
    mozhno_srazu — решать нечего, экран «Новая закупка» можно пропустить."""
    dokumenty = sobrat_documenty(put_papki, nastrojki, roli)
    zayavki = [d for d in dokumenty if d["rol"] == "zayavka"]
    oshibka = None
    rashozhdeniya: dict = {}
    if not zayavki:
        oshibka = "Заявка не найдена — добавьте файл заявки"
    else:
        try:
            karta = _prochitat_zayavku(zayavki[0], nastrojki, polzovatel_fio, dokumenty)
            rashozhdeniya = _sverit_otdelnye_fajly(dokumenty, karta, nastrojki)
        except ZayavkaNeProchitana as e:
            oshibka = str(e)
    neopredeleny = [
        d["imya"] for d in dokumenty
        if d["rol"] == "ne_ispolzuetsya" and Path(d["put"]).suffix.lower() in ROLI_CHITAEMYH
        and not rol_fajla.dokument_programmy(d["imya"]) and not d.get("oshibka_chteniya")
    ]
    return {
        "dokumenty": dokumenty,
        "oshibka_zayavki": oshibka,
        "neopredeleny": neopredeleny,
        "rashozhdeniya": rashozhdeniya,
        "mozhno_srazu": not oshibka and not neopredeleny and not rashozhdeniya,
    }


def najti_po_papke(nastrojki: Nastrojki, put_papki: str | Path) -> dict | None:
    """Карточка, уже созданная для этой папки (чтобы не заводить вторую)."""
    nuzhnaya = str(Path(put_papki))
    for s in spisok_zakupok(nastrojki):
        if str(Path(s.get("put_papki", ""))) == nuzhnaya and (s.get("karta") or {}).get("znacheniya"):
            return s
    return None


def _ubrat_pustye_kartochki(nastrojki: Nastrojki, put_papki: Path) -> None:
    """Карточки этой папки без данных заявки (остались от версии 0.5.3) — удалить."""
    import shutil
    for s in spisok_zakupok(nastrojki):
        if str(Path(s.get("put_papki", ""))) == str(put_papki) and not (s.get("karta") or {}).get("znacheniya"):
            shutil.rmtree(_put_zakupki(nastrojki, s["id"]), ignore_errors=True)


def _karta_v_slovar(karta) -> dict:
    return {
        "znacheniya": karta.znacheniya,
        "istochniki": {k: asdict(v) for k, v in karta.istochniki.items()},
        "pozicii": karta.pozicii,
        "preduprezhdeniya": karta.preduprezhdeniya,
    }


# ------------------------------------------------------------------ 1. opisanie_polej уже выше

# ------------------------------------------------------------------ 2. otkryt_kartochku

def otkryt_kartochku(
    nastrojki: Nastrojki,
    put_papki: str | Path,
    polzovatel_fio: str | None = None,
    id_zakupki: str | None = None,
    nomer_zakupki: str | None = None,
    roli: dict[str, str] | None = None,
) -> dict:
    """Создаёт карточку закупки в папке put_papki или открывает уже
    сохранённую (по id_zakupki). roli — роли файлов, выбранные сотрудником.
    Без прочитанной заявки карточка не создаётся — ZayavkaNeProchitana."""
    if id_zakupki:
        sostoyanie = _zagruzit(nastrojki, id_zakupki)
        if sostoyanie:
            return sostoyanie

    put_papki = Path(put_papki)
    dokumenty = sobrat_documenty(put_papki, nastrojki, roli)
    zayavki = [d for d in dokumenty if d["rol"] == "zayavka"]
    if not zayavki:
        raise ZayavkaNeProchitana("Заявка не найдена — добавьте файл заявки")
    karta = _prochitat_zayavku(zayavki[0], nastrojki, polzovatel_fio, dokumenty)
    _ubrat_pustye_kartochki(nastrojki, put_papki)

    id_zakupki = id_zakupki or uuid.uuid4().hex[:8]
    sostoyanie: dict[str, Any] = {
        "id": id_zakupki,
        "put_papki": str(put_papki),
        "nomer_zakupki": nomer_zakupki or nomer_iz_nazvaniya_papki(put_papki.name),
        # кто ведёт закупку — тот, кто завёл её в программе (экран «Кто вы»);
        # меняется только «Передать…»; ответственное лицо п. 1 на него не влияет
        "sotrudnik": polzovatel_fio,
        "sotrudnik_istochnik": "zavel",
        "versiya_chteniya": VERSIYA_CHTENIYA,
        "etap": "chernovik",
        "versiya": 1,
        "dokumenty": dokumenty,
        "forma_dogovora": None,
        "karta": None,
        "rashozhdeniya": {},
    }

    sostoyanie["karta"] = _karta_v_slovar(karta)
    sostoyanie["put_zayavki"] = zayavki[0]["put"]

    # отдельные файлы: сверка с заявкой (раздел 9 ТЗ)
    sostoyanie["rashozhdeniya"] = _sverit_otdelnye_fajly(dokumenty, karta, nastrojki)

    _sohranit(nastrojki, id_zakupki, sostoyanie)
    return sostoyanie


def _rashozhdenie_v_slovar(r) -> dict:
    return {"chto": r.chto, "v_zayavke": r.v_zayavke, "v_fajle": r.v_fajle}


def _sverit_otdelnye_fajly(dokumenty: list[dict], karta, nastrojki: Nastrojki | None = None) -> dict:
    rashozhdeniya: dict[str, list[dict]] = {}
    for d in dokumenty:
        if d.get("oshibka_chteniya"):
            continue
        put = _dlya_chteniya(d["put"], nastrojki) if nastrojki else d["put"]
        if d["rol"] == "raschet_nmck" and karta.znacheniya.get("nmck_raschet", {}).get("kp"):
            kp_iz_fajla = sverka.procitat_kp_iz_fajla_rascheta(put)
            r = sverka.sverit_nmck_kp(karta.znacheniya["nmck_raschet"]["kp"], kp_iz_fajla)
            if r:
                rashozhdeniya["raschet_nmck"] = [_rashozhdenie_v_slovar(x) for x in r]
        if d["rol"] == "tz" and karta.pozicii and Path(put).suffix.lower() == ".docx":
            r = sverka.sverit_tz(sverka.pozicii_zayavki_dlya_sverki(karta.pozicii), sverka.procitat_pozicii_tz(put))
            if r:
                rashozhdeniya["tz"] = [_rashozhdenie_v_slovar(x) for x in r]
        if d["rol"] == "instrukciya" and karta.znacheniya.get("instrukciya_tovara"):
            instr_iz_fajla = sverka.procitat_instrukciyu_iz_fajla(put)
            r = sverka.sverit_instrukciyu(karta.znacheniya["instrukciya_tovara"], instr_iz_fajla)
            if r:
                rashozhdeniya["instrukciya"] = [_rashozhdenie_v_slovar(x) for x in r]
    return rashozhdeniya


# ------------------------------------------------------------------ 3. zapolnit_iz_zayavki

def zapolnit_iz_zayavki(nastrojki: Nastrojki, id_zakupki: str, polzovatel_fio: str | None = None) -> dict:
    """Перечитывает заявку, заполняет пустые поля; решения сотрудника
    (площадка, форма договора, столбцы ТЗ) сохраняются."""
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    if sostoyanie is None:
        raise ValueError(f"закупка {id_zakupki} не найдена")

    put_zayavki = sostoyanie.get("put_zayavki")
    if not put_zayavki:
        return sostoyanie

    karta = _izvlech_iz_zayavki(Path(put_zayavki), nastrojki, polzovatel_fio, sostoyanie.get("dokumenty"))
    novaya = _karta_v_slovar(karta)

    # решения сотрудника не затираются: поля, которые он менял, остаются
    staraya = sostoyanie.get("karta") or {"znacheniya": {}, "istochniki": {}}
    izmeneno_vruchnuyu = {
        k for k, v in staraya.get("istochniki", {}).items()
        if v.get("fajl") == "sotrudnik" or v.get("yacheyka") == VYBOR_IZ_SPRAVOCHNIKA
    }
    for k in izmeneno_vruchnuyu:
        if k in staraya["znacheniya"]:
            novaya["znacheniya"][k] = staraya["znacheniya"][k]
            novaya["istochniki"][k] = staraya["istochniki"][k]
    if "ploshchadka" in staraya.get("znacheniya", {}):
        novaya["znacheniya"]["ploshchadka"] = staraya["znacheniya"]["ploshchadka"]

    sostoyanie["karta"] = novaya
    sostoyanie["versiya_chteniya"] = VERSIYA_CHTENIYA
    _sohranit(nastrojki, id_zakupki, sostoyanie)
    return sostoyanie


# ------------------------------------------------- снимок заявки и её изменения (раздел 8 ТЗ, этап 6)

_POLYA_PODPISI: dict[str, str] | None = None


def _podpis_polya(imya: str) -> str:
    global _POLYA_PODPISI
    if _POLYA_PODPISI is None:
        _POLYA_PODPISI = {p["imya"]: p["podpis"] for p in opisanie_polej()}
    return _POLYA_PODPISI.get(imya, imya)


def sohranit_snimok_zayavki(sostoyanie: dict, nastrojki: Nastrojki) -> None:
    """Снимок данных заявки (и позиций) на момент «Сформировать на
    согласование» — для сравнения «заявка изменилась» при повторном открытии
    закупки (раздел 8 ТЗ, этап 6)."""
    put_zayavki = sostoyanie.get("put_zayavki")
    if not put_zayavki:
        return
    svezhaya = _karta_v_slovar(_izvlech_iz_zayavki(Path(put_zayavki), nastrojki, None, sostoyanie.get("dokumenty")))
    sostoyanie["snimok_zayavki"] = {"znacheniya": svezhaya["znacheniya"], "pozicii": svezhaya["pozicii"]}


def najti_izmeneniya_zayavki(nastrojki: Nastrojki, id_zakupki: str) -> dict:
    """Сравнивает текущую заявку со снимком (раздел 8 ТЗ, этап 6). Пустой
    результат («est_izmeneniya»: False), если снимка нет (документ ещё не
    формировали) или заявка не менялась."""
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    if sostoyanie is None:
        raise ValueError(f"закупка {id_zakupki} не найдена")

    snimok = sostoyanie.get("snimok_zayavki")
    put_zayavki = sostoyanie.get("put_zayavki")
    if not snimok or not put_zayavki:
        return {"est_izmeneniya": False, "polya": [], "pozicii": [], "svezhaya_karta": None}

    svezhaya_karta = _karta_v_slovar(_izvlech_iz_zayavki(Path(put_zayavki), nastrojki, None, sostoyanie.get("dokumenty")))

    tekushchaya = sostoyanie.get("karta") or {"znacheniya": {}, "istochniki": {}}
    izmeneno_vruchnuyu = {
        k for k, v in tekushchaya.get("istochniki", {}).items() if v.get("fajl") == "sotrudnik"
    }

    staryye = snimok.get("znacheniya", {})
    novyye = svezhaya_karta.get("znacheniya", {})
    tekushchie = tekushchaya.get("znacheniya", {})
    polya = []
    for imya in sorted(set(staryye) | set(novyye)):
        if imya.startswith("_"):
            continue
        staroe, novoe = staryye.get(imya), novyye.get(imya)
        if staroe == novoe:
            continue
        konflikt = imya in izmeneno_vruchnuyu
        polya.append({
            "imya": imya, "podpis": _podpis_polya(imya),
            # при конфликте «было» — то, что вписал сотрудник (текущее значение
            # карточки), а не старый текст заявки из снимка — именно между
            # ними выбирает сотрудник
            "bylo": tekushchie.get(imya) if konflikt else staroe,
            "stalo": novoe,
            "konflikt": konflikt,
        })

    staryye_p = snimok.get("pozicii", [])
    novyye_p = svezhaya_karta.get("pozicii", [])
    pozicii = []
    for i in range(max(len(staryye_p), len(novyye_p))):
        stara_p = staryye_p[i] if i < len(staryye_p) else {}
        nova_p = novyye_p[i] if i < len(novyye_p) else {}
        for stolbec in sorted(set(stara_p) | set(nova_p)):
            sv = (stara_p.get(stolbec) or {}).get("znachenie")
            nv = (nova_p.get(stolbec) or {}).get("znachenie")
            if sv != nv:
                pozicii.append({"pozicija": i + 1, "stolbec": stolbec, "bylo": sv, "stalo": nv})

    return {
        "est_izmeneniya": bool(polya or pozicii),
        "polya": polya,
        "pozicii": pozicii,
        "svezhaya_karta": svezhaya_karta,
    }


def primenit_izmeneniya_zayavki(
    nastrojki: Nastrojki, id_zakupki: str, prinyat: bool, vybor_konfliktov: dict[str, str] | None = None,
) -> dict:
    """«Принять изменения» (prinyat=True) или «Оставить как было» (prinyat=
    False) по сравнению с заявкой. vybor_konfliktov — для полей, изменённых
    и вручную, и в заявке: {имя поля: "bylo"|"stalo"} (раздел 8 ТЗ, этап 6)."""
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    if sostoyanie is None:
        raise ValueError(f"закупка {id_zakupki} не найдена")

    izmeneniya = najti_izmeneniya_zayavki(nastrojki, id_zakupki)
    vybor_konfliktov = vybor_konfliktov or {}

    karta = sostoyanie.setdefault("karta", {"znacheniya": {}, "istochniki": {}, "pozicii": []})
    for p in izmeneniya["polya"]:
        imya = p["imya"]
        if p["konflikt"]:
            if vybor_konfliktov.get(imya) == "stalo":
                karta["znacheniya"][imya] = p["stalo"]
                karta["istochniki"].pop(imya, None)
            # иначе — оставляем значение сотрудника как есть (уже помечено «вручную»)
        elif prinyat:
            karta["znacheniya"][imya] = p["stalo"]
        else:
            # «Оставить как было» — значение не трогаем, но помечаем «вручную»,
            # чтобы следующее «Заполнить из заявки» его не затёрло молча
            karta["istochniki"][imya] = {"fajl": "sotrudnik", "list_": "", "yacheyka": "", "citata": "оставлено как было при изменении заявки"}

    if prinyat and izmeneniya["svezhaya_karta"]:
        karta["pozicii"] = izmeneniya["svezhaya_karta"]["pozicii"]

    if izmeneniya["svezhaya_karta"]:
        sostoyanie["snimok_zayavki"] = {
            "znacheniya": izmeneniya["svezhaya_karta"]["znacheniya"],
            "pozicii": izmeneniya["svezhaya_karta"]["pozicii"],
        }
    _sohranit(nastrojki, id_zakupki, sostoyanie)
    return sostoyanie


# ------------------------------------------------- сотрудники: кто ведёт, кто открыл

ZANYATOST = "zanyato.json"
SVEZHEST_ZANYATOSTI = 180  # секунд: открытая карточка подтверждает себя раз в минуту


def peredat_zakupku(nastrojki: Nastrojki, id_zakupki: str, fio: str) -> dict:
    """Руководитель передаёт закупку другому сотруднику."""
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    if sostoyanie is None:
        raise ValueError(f"закупка {id_zakupki} не найдена")
    sostoyanie["sotrudnik"] = fio
    sostoyanie["sotrudnik_istochnik"] = "peredano"
    _sohranit(nastrojki, id_zakupki, sostoyanie)
    osvobodit_kartochku(nastrojki, id_zakupki, None)
    return sostoyanie


def kto_otkryl(nastrojki: Nastrojki, id_zakupki: str, kto: dict) -> dict | None:
    """Кто другой сейчас редактирует карточку ({"fio", "komp"}) или None."""
    import time

    put = _put_zakupki(nastrojki, id_zakupki) / ZANYATOST
    try:
        zanyato = json.loads(put.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    svoya = zanyato.get("fio") == kto.get("fio") and zanyato.get("komp") == kto.get("komp")
    svezhaya = time.time() - float(zanyato.get("vremya", 0)) < SVEZHEST_ZANYATOSTI
    return None if svoya or not svezhaya else zanyato


def zanyat_kartochku(nastrojki: Nastrojki, id_zakupki: str, kto: dict) -> dict | None:
    """Отметить «карточка открыта у меня». Если её сейчас редактирует другой —
    ничего не меняется, возвращается он (тогда открыть только для просмотра)."""
    import time

    drugoj = kto_otkryl(nastrojki, id_zakupki, kto)
    if drugoj:
        return drugoj
    put = _put_zakupki(nastrojki, id_zakupki) / ZANYATOST
    try:
        put.write_text(json.dumps({"fio": kto.get("fio"), "komp": kto.get("komp"), "vremya": time.time()},
                                  ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass
    return None


def osvobodit_kartochku(nastrojki: Nastrojki, id_zakupki: str, kto: dict | None) -> None:
    """Снять отметку (kto=None — чья угодно, например при передаче закупки)."""
    put = _put_zakupki(nastrojki, id_zakupki) / ZANYATOST
    try:
        zanyato = json.loads(put.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return
    if kto is None or (zanyato.get("fio") == kto.get("fio") and zanyato.get("komp") == kto.get("komp")):
        put.unlink(missing_ok=True)


def zakupka_razmeshchena(nastrojki: Nastrojki, id_zakupki: str) -> None:
    """«Закупка размещена» — карточка удаляется из программы, документы в
    папке закупки (put_papki) не трогаются (раздел 14 ТЗ)."""
    import shutil
    shutil.rmtree(_put_zakupki(nastrojki, id_zakupki), ignore_errors=True)


# ------------------------------------------------------------------ 4. izmenit_pole

# источник «сотрудник выбрал из справочника» (п. 1) — как ручная правка, не затирается заявкой
VYBOR_IZ_SPRAVOCHNIKA = "выбор сотрудника"


def izmenit_pole(nastrojki: Nastrojki, id_zakupki: str, imya_polya: str, znachenie: Any) -> dict:
    """Меняет одно поле и пересчитывает зависимые (раздел 14 ТЗ)."""
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    if sostoyanie is None:
        raise ValueError(f"закупка {id_zakupki} не найдена")
    if not (imya_polya in ("forma_dogovora", "nomer_zakupki") or imya_polya.startswith("rol:")):
        _karta_ili_oshibka(sostoyanie)

    if imya_polya == "forma_dogovora":
        sostoyanie["forma_dogovora"] = znachenie
    elif imya_polya == "nomer_zakupki":
        sostoyanie["nomer_zakupki"] = znachenie
    elif imya_polya.startswith("rol:"):
        # смена роли файла сотрудником: "rol:<put>"
        put = imya_polya.split(":", 1)[1]
        for d in sostoyanie["dokumenty"]:
            if d["put"] == put:
                d["rol"] = znachenie
    elif re.match(r"^utverzhdayu\.(dolzhnost|fio)$", imya_polya):
        # УТВЕРЖДАЮ — выбор из справочника «Подписанты.xlsx» или вручную
        _, pole_vnutri = imya_polya.split(".")
        karta = sostoyanie.setdefault("karta", {"znacheniya": {}, "istochniki": {}, "pozicii": []})
        utv = karta["znacheniya"].setdefault("utverzhdayu", {"dolzhnost": "", "fio": ""})
        utv[pole_vnutri] = znachenie
        karta["istochniki"]["utverzhdayu"] = {"fajl": "sotrudnik", "list_": "", "yacheyka": "", "citata": "правка сотрудника"}
    elif imya_polya == "kontakt.sotrudnik":
        # п. 1: ответственное лицо — сотрудник из «Сотрудники отдела.xlsx»,
        # телефон подставляется из справочника («8» + городской + «, доб. » + добавочный)
        from izvlechenie_excel import zagruzit_sotrudnikov
        telefony = zagruzit_sotrudnikov(nastrojki.sp("Сотрудники отдела.xlsx"))
        if znachenie not in telefony:
            raise ValueError(f"сотрудника «{znachenie}» нет в справочнике «Сотрудники отдела.xlsx»")
        karta = sostoyanie.setdefault("karta", {"znacheniya": {}, "istochniki": {}, "pozicii": []})
        karta["znacheniya"]["kontakt_fio"] = znachenie
        karta["znacheniya"]["kontakt_tel"] = telefony[znachenie] or ""
        for pole in ("kontakt_fio", "kontakt_tel"):
            karta["istochniki"][pole] = {"fajl": "spravochnik", "list_": "Сотрудники отдела", "yacheyka": VYBOR_IZ_SPRAVOCHNIKA,
                                         "citata": f"справочник: {znachenie}"}
    elif imya_polya in ("opechatka.ispravit", "opechatka.ostavit"):
        # znachenie: {"najdeno": …, "ispravlenie": …} — слово исправляется (или
        # остаётся как в заявке) везде, где встречается
        karta = sostoyanie["karta"]
        najdeno = znachenie["najdeno"]
        if imya_polya == "opechatka.ispravit":
            mesta = _ispravit_slovo_vezde(karta, najdeno, znachenie["ispravlenie"])
            sostoyanie["soobshchenie"] = (
                f"Исправлено в {len(mesta)} {_mestah(len(mesta))}: {', '.join(mesta)}." if mesta
                else f"Слово «{najdeno}» не найдено — исправлять нечего."
            )
        else:
            sostoyanie.setdefault("opechatki_ostavleny", []).append(najdeno)
            sostoyanie["soobshchenie"] = f"«{najdeno}» оставлено как в заявке."
    elif imya_polya == "soglasovano.ubrat":
        karta = sostoyanie.setdefault("karta", {"znacheniya": {}, "istochniki": {}, "pozicii": []})
        spisok = karta["znacheniya"].setdefault("soglasovano", [])
        indeks = int(znachenie)
        if 0 <= indeks < len(spisok):
            spisok.pop(indeks)
    elif re.match(r"^soglasovano\.\d+\.(dolzhnost|fio)$", imya_polya):
        # «soglasovano.<индекс>.<поле>» — правка согласующего в карточке
        _, indeks, pole_vnutri = imya_polya.split(".")
        karta = sostoyanie.setdefault("karta", {"znacheniya": {}, "istochniki": {}, "pozicii": []})
        spisok = karta["znacheniya"].setdefault("soglasovano", [])
        indeks = int(indeks)
        while len(spisok) <= indeks:
            spisok.append({"dolzhnost": "", "fio": ""})
        spisok[indeks][pole_vnutri] = znachenie
        karta["istochniki"]["soglasovano"] = {"fajl": "sotrudnik", "list_": "", "yacheyka": "", "citata": "правка сотрудника"}
    elif imya_polya == "obesp_zayavki.ustanovleno":
        karta = sostoyanie.setdefault("karta", {"znacheniya": {}, "istochniki": {}, "pozicii": []})
        obesp = karta["znacheniya"].setdefault("obesp_zayavki", {})
        obesp["ustanovleno"] = (str(znachenie).lower() in ("true", "1", "установлено"))
        karta["istochniki"]["obesp_zayavki"] = {"fajl": "sotrudnik", "list_": "", "yacheyka": "", "citata": "правка сотрудника"}
    elif imya_polya == "nmck_raschet.sposob":
        # способ расчёта НМЦК (раздел 6 ТЗ): средняя / минимальная / смета
        karta = sostoyanie.setdefault("karta", {"znacheniya": {}, "istochniki": {}, "pozicii": []})
        raschet = karta["znacheniya"].setdefault("nmck_raschet", {})
        raschet["po_minimalnomu"] = (znachenie == "minimalnaya")
        raschet["smeta"] = (znachenie == "smeta")
        karta["istochniki"]["nmck_raschet"] = {"fajl": "sotrudnik", "list_": "", "yacheyka": "", "citata": "правка сотрудника"}
    elif re.match(r"^[a-z_]+\[\d+\]$", imya_polya):
        # «поле[номер]» — один абзац поля-списка (правка из экрана просмотра);
        # остальные абзацы, в том числе пустые, не трогаются
        imya, nomer = re.match(r"^([a-z_]+)\[(\d+)\]$", imya_polya).groups()
        karta = sostoyanie["karta"]
        spisok = list(karta["znacheniya"].get(imya) or [])
        if not 0 <= int(nomer) < len(spisok):
            raise ValueError(f"в поле {imya} нет абзаца № {int(nomer) + 1}")
        spisok[int(nomer)] = znachenie
        karta["znacheniya"][imya] = spisok
        karta["istochniki"][imya] = {"fajl": "sotrudnik", "list_": "", "yacheyka": "", "citata": "правка сотрудника"}
    elif imya_polya == "instrukciya_tovara":
        # правка инструкции целиком (кнопка «Править текст», раздел 4 ТЗ):
        # строка с переносами -> список абзацев, пустые не сохраняются
        karta = sostoyanie.setdefault("karta", {"znacheniya": {}, "istochniki": {}, "pozicii": []})
        abzacy = [a.strip() for a in str(znachenie).split("\n") if a.strip()]
        karta["znacheniya"]["instrukciya_tovara"] = abzacy
        karta["istochniki"]["instrukciya_tovara"] = {"fajl": "sotrudnik", "list_": "", "yacheyka": "", "citata": "правка сотрудника"}
    else:
        karta = sostoyanie.setdefault("karta", {"znacheniya": {}, "istochniki": {}, "pozicii": []})
        karta["znacheniya"][imya_polya] = znachenie
        karta["istochniki"][imya_polya] = {"fajl": "sotrudnik", "list_": "", "yacheyka": "", "citata": "правка сотрудника"}

        # зависимые пересчёты
        if imya_polya in ("obesp_isp_procent", "nmcd"):
            from raschety import chislo, summa_obespecheniya
            from izvlechenie_excel import format_chislo
            nmcd = karta["znacheniya"].get("nmcd")
            procent = karta["znacheniya"].get("obesp_isp_procent")
            try:
                summa = summa_obespecheniya(chislo(nmcd), chislo(procent)) if procent and nmcd else None
            except (ArithmeticError, ValueError, TypeError):
                summa = None  # вписано не число — сумма не считается, значение сотрудника сохраняется
            karta["znacheniya"]["obesp_isp_summa"] = format_chislo(summa) if summa is not None else ""

    _sohranit(nastrojki, id_zakupki, sostoyanie)
    return sostoyanie


# ------------------------------------------------------------------ 5. proverit

def proverit(nastrojki: Nastrojki, id_zakupki: str) -> dict:
    """Список ошибок и предупреждений самопроверки (раздел 11 ТЗ) + список
    «Решить» — то, что сотрудник должен выбрать сам."""
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    if sostoyanie is None:
        raise ValueError(f"закупка {id_zakupki} не найдена")

    karta_dict = sostoyanie.get("karta") or {}
    znacheniya = karta_dict.get("znacheniya", {})

    # «Решить» — коротко для строки над карточкой; те же пункты идут и в
    # самопроверку полными фразами со ссылкой «к полю» (раздел 11 ТЗ —
    # самопроверка и список «Решить» должны совпадать)
    reshit: list[dict] = []
    if not znacheniya.get("ploshchadka"):
        reshit.append({"pole": "ploshchadka", "tekst": "площадка", "fraza": "Не выбрана электронная площадка", "yakor": "s2"})
    if not sostoyanie.get("forma_dogovora"):
        reshit.append({"pole": "forma_dogovora", "tekst": "форма договора", "fraza": "Не подтверждена форма договора", "yakor": "s10"})
    soglasovano = znacheniya.get("soglasovano") or []
    if znacheniya and not soglasovano:
        # подразделения заявки нет в справочнике «подразделение → согласующий»
        reshit.append({
            "pole": "soglasovano.0.fio", "tekst": "согласующий",
            "fraza": "Не определён согласующий: подразделения заявки нет в справочнике — выберите или впишите вручную",
            "yakor": "s1",
        })
    for i, s in enumerate(soglasovano):
        if not s.get("fio"):
            reshit.append({
                "pole": f"soglasovano.{i}.fio", "tekst": f"ФИО согласующего № {i + 1}",
                "fraza": f"Не указаны И.О. Фамилия согласующего № {i + 1}", "yakor": "s1",
            })
    if not sostoyanie.get("nomer_zakupki"):
        reshit.append({"pole": "nomer_zakupki", "tekst": "номер закупки", "fraza": "Не указан номер закупки", "yakor": "s11"})
    # предмет — в «Решить» только если его нет в заявке (ниже, OBYAZATELNYE_IZ_ZAYAVKI)
    # или в нём возможная опечатка; прочитанный без замечаний — не спрашивается
    # (решение специалиста от 11.10.2026: никаких обязательных отметок)
    opechatki = opechatki_zayavki(nastrojki, sostoyanie) if znacheniya else []
    for o in opechatki:
        if "predmet" in o["gruppy"]:
            reshit.append({
                "pole": "opechatka.predmet", "tekst": "предмет закупки",
                "fraza": (f"Возможна опечатка в предмете закупки (п. 6): «{o['najdeno']}» → «{o['ispravlenie']}» — "
                          "нажмите «Исправить» или «Оставить как в заявке»"),
                "yakor": "predmet",
            })
            break
    # поля документа, которых не нашлось в заявке (другая форма заявки) —
    # сотрудник заполняет в карточке; без них документ не собирается
    for pole, podpis, yakor in OBYAZATELNYE_IZ_ZAYAVKI:
        if znacheniya and not znacheniya.get(pole):
            reshit.append({"pole": pole, "tekst": podpis, "fraza": f"Не найдено в заявке: {podpis} — заполните в карточке", "yakor": yakor})

    PODPISI_RASHOZHDENIJ = {"raschet_nmck": "Расчёт НМЦК", "instrukciya": "Инструкция по описанию товара", "tz": "Техническое задание"}
    for chto, r in (sostoyanie.get("rashozhdeniya") or {}).items():
        yakor = {"raschet_nmck": "s6", "tz": "s3"}.get(chto, "s8")
        podpis = PODPISI_RASHOZHDENIJ.get(chto, chto)
        reshit.append({
            "pole": f"rashozhdenie.{chto}", "tekst": f"{podpis} — расхождение с файлом",
            "fraza": f"{podpis}: расхождение с отдельным файлом — {len(r)} шт., исправьте заявку или файл", "yakor": yakor,
        })

    gotovo: list[str] = []
    preduprezhdeniya: list[dict] = []
    raschet = znacheniya.get("nmck_raschet") or {}
    otklonenie_nmcd = None
    if raschet.get("srednyaya") and znacheniya.get("nmcd"):
        from raschety import chislo
        otklonenie_nmcd = chislo(znacheniya["nmcd"]) - chislo(raschet["srednyaya"])
    otklonenie_nmcd_tekst = format_chislo(abs(otklonenie_nmcd)) if otklonenie_nmcd else None
    if otklonenie_nmcd:
        storona = "ниже" if otklonenie_nmcd < 0 else "выше"
        fraza_otklonenie = f"НМЦД {storona} средней по КП на {otklonenie_nmcd_tekst} руб. — нужно посмотреть и подтвердить просмотр"
        if znacheniya.get("nmcd_otklonenie_prosmotreno"):
            gotovo.append(f"расхождение {otklonenie_nmcd_tekst} руб. просмотрено")
            preduprezhdeniya.append({"tekst": f"НМЦД {storona} средней по КП на {otklonenie_nmcd_tekst} руб. — просмотрено", "yakor": "s6"})
        else:
            reshit.append({"pole": "nmcd_otklonenie_prosmotreno", "tekst": "расхождение НМЦД со средней", "fraza": fraza_otklonenie, "yakor": "s6"})
    # опечатки вне предмета (нацрежим, позиции ТЗ) — предупреждение, не блокирует
    for o in opechatki:
        if "predmet" in o["gruppy"]:
            continue
        preduprezhdeniya.append({
            "tekst": f"Возможна опечатка: «{o['najdeno']}» → «{o['ispravlenie']}» ({', '.join(o['mesta'])})",
            "yakor": "nacrezhim" if o["gruppy"][0] == "nacrezhim" else "opechatki-pozicij",
        })

    # полнота документа: пустые приложения, ИТОГО и сумма позиций ≠ НМЦД
    reshit.extend(proverit_polnotu(nastrojki, sostoyanie))

    oshibki: list[dict] = [{"tekst": r["fraza"], "yakor": r["yakor"]} for r in reshit]
    preduprezhdeniya.extend(preduprezhdeniya_zayavki(nastrojki, sostoyanie))

    if znacheniya:
        class _K:
            pass
        k = _K()
        k.znacheniya = znacheniya
        k.pozicii = karta_dict.get("pozicii", [])
        k.preduprezhdeniya = karta_dict.get("preduprezhdeniya", [])
        slovar_put = nastrojki.sp("Словарь для проверки опечаток.xlsx")
        slovar = zagruzit_slovar(slovar_put) if slovar_put.exists() else set()
        r = proverit_kartochku(k, slovar)
        oshibki.extend({"tekst": t, "yakor": "s6"} for t in r.oshibki)
        preduprezhdeniya.extend({"tekst": t, "yakor": "s6"} for t in r.preduprezhdeniya)

    # «можно сформировать» — нет пунктов «Решить» и нет остальных ошибок
    # самопроверки (пункты «Решить» в oshibki уже учтены выше)
    oshibki_krome_reshit = [o for o in oshibki if o["tekst"] not in {r["fraza"] for r in reshit}]
    rezultat = {
        "reshit": reshit,
        "vsego_reshaemyh": len(reshit) + len(gotovo),
        "gotovo": gotovo,
        "oshibki": oshibki,
        "preduprezhdeniya": preduprezhdeniya,
        "mozhno_sformirovat": not reshit and not oshibki_krome_reshit,
        "otklonenie_nmcd": otklonenie_nmcd_tekst,
        "otklonenie_nmcd_storona": ("ниже" if otklonenie_nmcd < 0 else "выше") if otklonenie_nmcd else None,
        "opechatki": opechatki,
    }
    sostoyanie["proverka"] = rezultat
    _sohranit(nastrojki, id_zakupki, sostoyanie)
    return rezultat


# ------------------------------------------------- самопроверка полноты документа

def _chislo_ili_none(tekst):
    from raschety import chislo
    try:
        return chislo(tekst) if str(tekst).strip() else None
    except (ArithmeticError, ValueError, TypeError):
        return None


def _stolbec_pozicii(poziciya: dict, *fragmenty: str):
    for nazvanie, yach in poziciya.items():
        nizhnij = " ".join(nazvanie.lower().split())
        if any(f in nizhnij for f in fragmenty) and "сумм" not in nizhnij:
            return yach.get("znachenie") if isinstance(yach, dict) else yach
    return None


def _strok_pozicij_v_tablice(tablica) -> int:
    """Строки позиций таблицы файла ТЗ — те, где в первой ячейке номер."""
    return sum(1 for r in tablica.rows if re.fullmatch(r"\d+\.?", r.cells[0].text.strip()))


def proverit_polnotu(nastrojki: Nastrojki, sostoyanie: dict) -> list[dict]:
    """Части документа, которые получились бы пустыми или неверными (раздел 11
    ТЗ): Приложение № 2 без расчёта НМЦК или с ИТОГО ≠ НМЦД, Приложение № 3 без
    позиций или без таблицы характеристик из файла ТЗ, сумма позиций ≠ НМЦД.
    Каждый пункт — в «Решить»: документ с пустыми частями не собирается."""
    from decimal import Decimal

    karta = sostoyanie.get("karta") or {}
    znach = karta.get("znacheniya") or {}
    if not znach:
        return []
    pozicii = karta.get("pozicii") or []
    nmcd = _chislo_ili_none(znach.get("nmcd") or "")
    reshit: list[dict] = []

    def dobavit(pole, tekst, fraza, yakor="s6"):
        reshit.append({"pole": pole, "tekst": tekst, "fraza": fraza, "yakor": yakor})

    # Приложение № 2 — расчёт НМЦК
    raschet = znach.get("nmck_raschet") or {}
    itogo = None
    if not raschet.get("smeta"):
        if raschet.get("po_obshchej_summe") and raschet.get("kp"):
            pass  # таблица шаблона; НМЦД сверяется со средней (расхождение — выше)
        elif znach.get("_nmck_tarif") and znach["_nmck_tarif"].get("pozicii"):
            itogo = znach["_nmck_tarif"].get("itogo")
        elif znach.get("_nmck_poz") and znach["_nmck_poz"].get("stroki"):
            itogo = znach["_nmck_poz"].get("itogo")
            v_raschete = len(znach["_nmck_poz"]["stroki"])
            if pozicii and v_raschete != len(pozicii):
                dobavit("polnota.prilozhenie_2", "приложение № 2: число позиций",
                        f"Приложение № 2: в расчёте НМЦК {v_raschete} поз., в таблице позиций заявки {len(pozicii)} — "
                        "проверьте лист «Расчет НМЦК»")
        else:
            dobavit("polnota.prilozhenie_2", "приложение № 2 пустое",
                    "Приложение № 2 пустое: не найден расчёт НМЦК в заявке — добавьте файл расчёта "
                    "или проверьте лист «Расчет НМЦК»")
        itogo_chislo = _chislo_ili_none(itogo) if itogo else None
        if itogo is not None and nmcd is not None and itogo_chislo != nmcd:
            dobavit("polnota.itogo", "ИТОГО приложения № 2 ≠ НМЦД",
                    f"Приложение № 2: ИТОГО расчёта НМЦК {itogo} руб. не равно НМЦД в п. 8 {znach.get('nmcd')} руб. — "
                    "проверьте лист «Расчет НМЦК» и п. 8")

    # Приложение № 3 — позиции и таблица характеристик
    if not pozicii:
        dobavit("polnota.prilozhenie_3", "приложение № 3 пустое",
                "Приложение № 3 пустое: не найдена таблица позиций в заявке — проверьте лист «Заявка» "
                "(строка заголовков с «Ед. изм.»)", "s3")
    else:
        summa = Decimal(0)
        for p in pozicii:
            kol = _chislo_ili_none(_stolbec_pozicii(p, "кол-во", "количество") or "")
            cena = _chislo_ili_none(_stolbec_pozicii(p, "цена") or "")
            if kol is None or cena is None:
                summa = None
                break
            summa += kol * cena
        if summa is None:
            dobavit("polnota.pozicii", "цены и количество позиций",
                    "Приложение № 3: не у всех позиций прочитаны количество и цена — проверьте таблицу позиций заявки", "s3")
        elif nmcd is not None and abs(summa.quantize(Decimal("0.01")) - nmcd) > Decimal("0.01") * len(pozicii):
            dobavit("polnota.summa", "сумма позиций ≠ НМЦД",
                    f"Сумма позиций приложения № 3 ({format_chislo(summa)} руб.) не равна НМЦД в п. 8 "
                    f"({znach.get('nmcd')} руб.) — проверьте цены и количество", "s3")
    from docx import Document
    for put in fajly_tz_zakupki(nastrojki, sostoyanie):
        imya = Path(put).name
        try:
            tablicy = Document(put).tables
        except Exception:
            tablicy = []
        v_tablice = sum(_strok_pozicij_v_tablice(t) for t in tablicy)
        if not tablicy:
            dobavit("polnota.tz", "таблица характеристик из файла ТЗ",
                    f"Приложение № 3: в файле ТЗ «{imya}» не найдена таблица характеристик — проверьте файл", "s3")
        elif pozicii and v_tablice != len(pozicii):
            dobavit("polnota.tz", "таблица характеристик: число позиций",
                    f"Приложение № 3: в таблице характеристик файла ТЗ «{imya}» {v_tablice} поз., "
                    f"в заявке {len(pozicii)} — проверьте файл", "s3")
    return reshit


# ------------------------------------- предупреждения по заявке (не блокируют)

def _kp_nomera_i_daty(znach: dict) -> list[tuple[int, date, str]]:
    """Входящие номера и даты КП: из таблицы расчёта «по позициям» или из
    списка КП расчёта по общей сумме."""
    kp: list[tuple[int, date, str]] = []
    teksty = []
    nmck_poz = znach.get("_nmck_poz") or {}
    teksty += [s for s in (nmck_poz.get("niz") or []) if s]
    teksty += [f"№ {k.get('vh_nomer')} от {k.get('vh_data')}" for k in (znach.get("nmck_raschet") or {}).get("kp") or []]
    for s in teksty:
        m = re.search(r"№\s*(\d+)\S*\s+от\s+(\d{1,2})\.(\d{2})\.(\d{4})", s)
        if not m:
            continue
        try:
            kp.append((int(m.group(1)), date(int(m.group(4)), int(m.group(3)), int(m.group(2))), f"№ {m.group(1)} от {m.group(2)}.{m.group(3)}.{m.group(4)}"))
        except ValueError:
            continue
    return kp


STEPENI = re.compile(r"·\s?10\d+|(?<![\w])(?:см|мм|дм|м)[23](?!\d)")


def vozmozhno_poteryany_stepeni(fajly_tz: list[Path]) -> list[str]:
    """Места таблиц файла ТЗ, где степень набрана обычными цифрами («5·1016»,
    «1/см3»): если верхние индексы в файле есть — они переносятся как есть;
    если нет — показать, где проверить."""
    from docx import Document
    from docx.oxml.ns import qn

    mesta: list[str] = []
    for put in fajly_tz:
        for tablica in Document(put).tables:
            for stroka in tablica.rows:
                nomer = stroka.cells[0].text.strip().rstrip(".")
                najdeno: list[str] = []
                for c in stroka.cells:
                    for r in c._tc.iter(qn("w:r")):
                        rpr = r.find(qn("w:rPr"))
                        if rpr is not None and rpr.find(qn("w:vertAlign")) is not None:
                            continue
                        tekst = "".join(t.text or "" for t in r.iter(qn("w:t")))
                        najdeno += [m.group(0).replace(" ", "") for m in STEPENI.finditer(tekst)]
                if najdeno:
                    gde = f"позиция {nomer}" if nomer.isdigit() else "таблица"
                    mesta.append(f"{Path(put).name}, {gde}: {', '.join(dict.fromkeys(najdeno))}")
    return mesta


def preduprezhdeniya_zayavki(nastrojki: Nastrojki, sostoyanie: dict) -> list[dict]:
    """Замечания к заявке, которые не мешают собрать документ (раздел 11 ТЗ):
    согласующего нет в листе согласования заявки, исправленная дата КП,
    входящие номера КП не по порядку дат, возможно потерянные степени."""
    znach = (sostoyanie.get("karta") or {}).get("znacheniya") or {}
    preduprezhdeniya: list[dict] = []
    list_soglasovaniya = znach.get("_list_soglasovaniya")
    if list_soglasovaniya:
        tekst_lista = " ".join(list_soglasovaniya)
        for s in znach.get("soglasovano") or []:
            fio = (s.get("fio") or "").strip()
            # фамилия — слово без точек: «О.С. Волкова», «Свистунов А.А.»
            familiya = next((w for w in fio.replace(".", ". ").split() if len(w) > 2 and "." not in w), "")
            if familiya and familiya not in tekst_lista:
                preduprezhdeniya.append({"tekst": f"{fio} нет в листе согласования заявки", "yakor": "s1"})
    for d in (znach.get("_nmck_poz") or {}).get("ispravleny_daty") or []:
        preduprezhdeniya.append({
            "tekst": (f"Дата КП «{d['bylo']}» (лист «{d['list']}», {d['yacheyka']}) исправлена на «{d['stalo']}» "
                      "в приложении № 2 — проверьте заявку"),
            "yakor": "s6",
        })
    kp = _kp_nomera_i_daty(znach)
    po_nomeram = sorted(kp)
    if any(a[1] > b[1] for a, b in zip(po_nomeram, po_nomeram[1:])):
        preduprezhdeniya.append({
            "tekst": ("Входящие номера КП не по порядку дат: " + "; ".join(k[2] for k in po_nomeram)
                      + " — проверьте номера и даты КП"),
            "yakor": "s6",
        })
    mesta = vozmozhno_poteryany_stepeni(fajly_tz_zakupki(nastrojki, sostoyanie))
    if mesta:
        preduprezhdeniya.append({
            "tekst": "Таблица характеристик: возможно, потеряны степени, проверьте — " + "; ".join(mesta),
            "yakor": "s3",
        })
    return preduprezhdeniya


METKI_NE_ZAPOLNENY = re.compile(r"\{\{|\{%|\}\}|%\}|(?<!\w)None(?!\w)")


def proverit_dokument(put: str | Path, sostoyanie: dict, fajly_tz: list[Path]) -> list[str]:
    """Собранный документ перед просмотром и сохранением: приложения не пустые,
    не осталось меток шаблона и «None». Возвращает тексты ошибок."""
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(put)
    oshibki: list[str] = []
    chasti: dict[str, dict] = {}
    tekushchaya = None
    for el in doc.element.body.iterchildren():
        if el.tag == qn("w:p"):
            tekst = Paragraph(el, doc).text
            m = re.match(r"\s*Приложение\s*№\s*(\d)\s*$", tekst)
            if m:
                tekushchaya = m.group(1)
                chasti[tekushchaya] = {"tablic": 0, "abzacy": []}
            elif tekushchaya and tekst.strip():
                chasti[tekushchaya]["abzacy"].append(tekst)
            if METKI_NE_ZAPOLNENY.search(tekst):
                oshibki.append(f"В документе осталась незаполненная метка или «None»: «{tekst.strip()[:80]}»")
        elif el.tag == qn("w:tbl"):
            if tekushchaya:
                chasti[tekushchaya]["tablic"] += 1
            for r in Table(el, doc).rows:
                for c in r.cells:
                    if METKI_NE_ZAPOLNENY.search(c.text):
                        oshibki.append(f"В таблице документа осталась незаполненная метка или «None»: «{c.text.strip()[:80]}»")
                        break
    raschet = ((sostoyanie.get("karta") or {}).get("znacheniya") or {}).get("nmck_raschet") or {}
    if not raschet.get("smeta") and not chasti.get("2", {}).get("tablic"):
        oshibki.append("Приложение № 2 пустое: в документе нет таблицы расчёта НМЦК")
    nuzhno_tablic = 1 + sum(len(Document(f).tables) for f in fajly_tz)
    v_3 = chasti.get("3", {}).get("tablic", 0)
    if v_3 < nuzhno_tablic:
        oshibki.append("Приложение № 3 неполное: нет таблицы позиций" if not v_3
                       else "Приложение № 3 неполное: нет таблицы характеристик из файла ТЗ")
    abzacy_4 = chasti.get("4", {}).get("abzacy", [])
    nachalo = next((i for i, a in enumerate(abzacy_4) if "Инструкция по описанию товара" in a), None)
    if nachalo is None or not any(a.strip() for a in abzacy_4[nachalo + 1:]):
        oshibki.append("Приложение № 4: инструкция по описанию товара пустая")
    return list(dict.fromkeys(oshibki))


# ------------------------------------------------------------------ 6. sformirovat

def _svobodnoe_imya(put: Path) -> Path:
    """«Сохранить как новый файл»: имя с (2), (3)… — старый файл остаётся."""
    if not put.exists():
        return put
    n = 2
    while (put.parent / f"{put.stem} ({n}){put.suffix}").exists():
        n += 1
    return put.parent / f"{put.stem} ({n}){put.suffix}"


def fajly_tz_zakupki(nastrojki: Nastrojki, sostoyanie: dict) -> list[Path]:
    """Отдельные файлы ТЗ закупки (.docx или переведённые в .docx) — их
    таблицы характеристик идут в Приложение № 3 целиком."""
    fajly = []
    for d in sostoyanie.get("dokumenty") or []:
        if d.get("rol") == "tz" and not d.get("oshibka_chteniya"):
            put = _dlya_chteniya(d["put"], nastrojki)
            if Path(put).suffix.lower() == ".docx" and Path(put).exists():
                fajly.append(Path(put))
    return fajly


def podgotovit_prosmotr(
    nastrojki: Nastrojki,
    id_zakupki: str,
    etap: str = "soglasovanie",
    d_razmeshcheniya: date | None = None,
    dazhe_s_oshibkami: bool = False,
) -> dict:
    """Собирает извещение (и на этапе согласования — проект договора) во
    временную папку карточки — для экрана просмотра. В папку закупки ничего
    не пишется: сохраняет ровно эти файлы sohranit_prosmotr().
    dazhe_s_oshibkami — собрать и при ошибках самопроверки (экран просмотра
    показывает их после правки; сохранить при ошибках нельзя)."""
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    if sostoyanie is None:
        raise ValueError(f"закупка {id_zakupki} не найдена")
    _karta_ili_oshibka(sostoyanie)

    proverka = proverit(nastrojki, id_zakupki)
    if not proverka["mozhno_sformirovat"] and not dazhe_s_oshibkami and etap != "razmeshchenie_cherez_proverku_propustit":
        return {"ok": False, "oshibka": "не всё решено или есть ошибки самопроверки", "proverka": proverka}
    sostoyanie = _zagruzit(nastrojki, id_zakupki)

    karta = sostoyanie["karta"]
    znacheniya = dict(karta["znacheniya"])
    nomer = sostoyanie["nomer_zakupki"]
    nomer_fajl = re.sub(r"[\\/:*?\"<>|]", "-", nomer)

    if etap == "soglasovanie":
        god = str(datetime.now().year)
        znacheniya["god"] = god
        for imya in ("d_razmeshcheniya", "d_zaprosy_do", "d_podacha_do", "d_rassmotrenie", "d_aukcion", "d_itogi"):
            znacheniya[imya] = f"__.__.{god}г."
        znacheniya["t_podacha_do"] = "09 ч. 00 мин."
        znacheniya["t_aukcion"] = "11 ч. 00 мин."
        imya_izveshcheniya = f"Аукцион. Извещение № {nomer_fajl} (на согласование).docx"
    else:
        kal = zagruzit_kalendar(nastrojki.sp("Производственный календарь.xlsx"))
        d = d_razmeshcheniya or date.today()
        try:
            daty = rasschitat_daty(d, kal)
            d_podacha_override = sostoyanie.get("d_podacha_do_vybrano")
            if d_podacha_override:
                d_podacha_do = date.fromisoformat(d_podacha_override)
                ost = rasschitat_ostalnye_daty(d_podacha_do, kal)
                daty = DatyIzveshcheniya(daty.d_razmeshcheniya, ost.d_zaprosy_do, d_podacha_do, ost.d_rassmotrenie, ost.d_aukcion, ost.d_itogi)
        except (NetKalendaryaNaGod, DataRazmeshcheniyaNerabochaya) as e:
            return {"ok": False, "oshibka": str(e)}
        t_podacha = sostoyanie.get("t_podacha_vybrano") or "09:00"
        t_aukcion = sostoyanie.get("t_aukcion_vybrano") or "11:00"
        znacheniya["god"] = str(d.year)
        znacheniya.update(format_dannye_dlya_shablona(daty, t_podacha, t_aukcion))
        imya_izveshcheniya = f"Аукцион. Извещение № {nomer_fajl}.docx"
    imya_dogovora = f"Аукцион. Проект договора № {nomer_fajl}.docx" if etap == "soglasovanie" else None

    papka_prosmotra = _put_zakupki(nastrojki, id_zakupki) / "_prosmotr" / etap
    if papka_prosmotra.exists():
        import shutil
        shutil.rmtree(papka_prosmotra, ignore_errors=True)
    papka_prosmotra.mkdir(parents=True, exist_ok=True)
    put_izveshcheniya = papka_prosmotra / imya_izveshcheniya
    put_dogovora = papka_prosmotra / imya_dogovora if imya_dogovora else None

    put_shablona = nastrojki.put_dannyh / "shablony" / "zakrytyj_aukcion.docx"
    vremennaya = _put_zakupki(nastrojki, id_zakupki) / "_vrem"
    tz_zagolovki, tz_stroki, _, _ = postroit_tablicu_tz(nastrojki, sostoyanie)
    fajly_tz = fajly_tz_zakupki(nastrojki, sostoyanie)
    sborka.sobrat_izveshchenie(
        put_shablona, znacheniya, karta["pozicii"], put_izveshcheniya, vremennaya,
        tablica_tz_gotovaya=(tz_zagolovki, tz_stroki),
        fajly_tz=fajly_tz,
        ispravleniya=karta.get("ispravleniya_opechatok"),
    )
    if put_dogovora is not None:
        formy = sborka.zagruzit_formy_dogovora(nastrojki.sp(SPRAVOCHNIK_FORM_DOGOVORA))
        forma = next((f for f in formy if f["nazvanie"] == sostoyanie.get("forma_dogovora")), None) \
            or sborka.podskazat_formu_dogovora(znacheniya.get("predmet_im", ""), formy)
        sborka.sobrat_proekt_dogovora(
            nastrojki.put_dannyh / "dogovory" / forma["fajl"], znacheniya, put_dogovora, zaglushka=forma["zaglushka"],
        )

    prosmotr = {
        "ok": True, "etap": etap,
        "izveshchenie": str(put_izveshcheniya), "dogovor": str(put_dogovora) if put_dogovora else None,
        "imya_izveshcheniya": imya_izveshcheniya, "imya_dogovora": imya_dogovora,
        "papka": sostoyanie["put_papki"],
    }
    # собранный документ проверяется перед просмотром: пустые приложения,
    # оставшиеся метки — ошибка, сохранить нельзя
    oshibki_dokumenta = proverit_dokument(put_izveshcheniya, sostoyanie, fajly_tz)
    prosmotr["oshibki_dokumenta"] = oshibki_dokumenta
    if oshibki_dokumenta:
        proverka = dict(proverka, mozhno_sformirovat=False,
                        oshibki=proverka["oshibki"] + [{"tekst": o, "yakor": "s1"} for o in oshibki_dokumenta])
    sostoyanie["prosmotr"] = prosmotr
    _sohranit(nastrojki, id_zakupki, sostoyanie)
    return dict(prosmotr, proverka=proverka)


def sohranit_prosmotr(nastrojki: Nastrojki, id_zakupki: str, zamena: str | None = None) -> dict:
    """«Сохранить в папку» на экране просмотра: копирует в папку закупки
    ровно те файлы, что показаны. zamena: None | "zamenit" | "novyj" | "otmena"
    — ответ на окно «файл уже есть в папке» (одинаково для извещения и договора)."""
    import shutil

    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    if sostoyanie is None:
        raise ValueError(f"закупка {id_zakupki} не найдена")
    prosmotr = sostoyanie.get("prosmotr") or {}
    gotovye = [prosmotr.get("izveshchenie"), prosmotr.get("dogovor")]
    if not prosmotr.get("izveshchenie") or not all(Path(p).exists() for p in gotovye if p):
        return {"ok": False, "oshibka": "Документ не подготовлен — нажмите «Сформировать» и проверьте его на экране."}
    etap = prosmotr["etap"]
    # правки из просмотра проверены самопроверкой — с ошибками не сохраняем
    proverka = proverit(nastrojki, id_zakupki)
    if not proverka["mozhno_sformirovat"]:
        return {"ok": False, "oshibka": "Сначала исправьте замечания самопроверки — они показаны над документом.",
                "proverka": proverka}
    if prosmotr.get("oshibki_dokumenta"):
        return {"ok": False, "oshibka": "Документ неполный: " + "; ".join(prosmotr["oshibki_dokumenta"]),
                "proverka": proverka}
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    put_papki = Path(sostoyanie["put_papki"])

    put_izveshcheniya = put_papki / prosmotr["imya_izveshcheniya"]
    put_dogovora = put_papki / prosmotr["imya_dogovora"] if prosmotr.get("imya_dogovora") else None
    uzhe_est = [p for p in (put_izveshcheniya, put_dogovora) if p is not None and p.exists()]
    if uzhe_est and zamena is None:
        # файл правили в Word после того, как программа его сохранила — при замене правки пропадут
        sohraneno = sostoyanie.get("sohraneno_programmoj") or {}
        pravili = [p.name for p in uzhe_est if str(p) in sohraneno and p.stat().st_mtime_ns != sohraneno[str(p)]]
        return {"ok": False, "nuzhno_reshenie_fajl": uzhe_est[0].name, "etap": etap, "pravili_v_word": pravili}
    if uzhe_est and zamena == "otmena":
        return {"ok": False, "otmeneno": True}
    if zamena == "novyj":
        put_izveshcheniya = _svobodnoe_imya(put_izveshcheniya)
        put_dogovora = _svobodnoe_imya(put_dogovora) if put_dogovora is not None else None

    shutil.copyfile(prosmotr["izveshchenie"], put_izveshcheniya)
    rezultat = {"ok": True, "etap": etap, "izveshchenie": str(put_izveshcheniya)}
    if put_dogovora is not None:
        shutil.copyfile(prosmotr["dogovor"], put_dogovora)
        rezultat["dogovor"] = str(put_dogovora)

    # время сохранения программой — чтобы потом заметить правки в Word
    sohraneno = sostoyanie.setdefault("sohraneno_programmoj", {})
    for p in (put_izveshcheniya, put_dogovora):
        if p is not None:
            sohraneno[str(p)] = p.stat().st_mtime_ns
    if etap == "soglasovanie":
        sostoyanie["etap"] = "na_soglasovanii"
        sostoyanie["data_etapa"] = date.today().isoformat()
        sohranit_snimok_zayavki(sostoyanie, nastrojki)
    else:
        sostoyanie["etap"] = "razmeshchenie"
    sostoyanie["poslednie_fajly"] = rezultat
    _sohranit(nastrojki, id_zakupki, sostoyanie)
    return rezultat


def sformirovat(
    nastrojki: Nastrojki,
    id_zakupki: str,
    etap: str = "soglasovanie",
    d_razmeshcheniya: date | None = None,
    zamena: str | None = None,
) -> dict:
    """Собирает извещение (и на этапе согласования — проект договора) и сразу
    сохраняет в папку закупки (для программы и тестов; сотрудник сохраняет через
    экран просмотра: podgotovit_prosmotr → sohranit_prosmotr). zamena: None |
    "zamenit" | "novyj" | "otmena" — ответ на окно «файл уже есть в папке»."""
    prosmotr = podgotovit_prosmotr(nastrojki, id_zakupki, etap, d_razmeshcheniya)
    if not prosmotr.get("ok"):
        return prosmotr
    return sohranit_prosmotr(nastrojki, id_zakupki, zamena)

# ------------------------------------------------------------------ таблица ТЗ

def _sluzhebnye_stolbcy_zayavki(pozicii: list[dict]) -> list[str]:
    """Столбцы заявки, которые по умолчанию не идут в документ (раздел 5 ТЗ:
    «Общая сумма», НКМИ, ОКПД2, ФФ ОМС и т. п.) — сотрудник может их вернуть
    на экране «Таблица ТЗ». Столбец меры по ПП № 1875 сюда не входит — у него
    особое правило (последний, нельзя убрать)."""
    if not pozicii:
        return []
    return [n for n in pozicii[0].keys()
            if "1875" not in n and (tablica_tz._sluzhebnyj(n) or tablica_tz.pustoj_stolbec(pozicii, n))]


# ------------------------------------------- дочитать недостающее из заявки

# что программа умеет читать из заявки; растёт с версиями — карточка, созданная
# старой версией, дочитывает недостающее один раз при открытии
VERSIYA_CHTENIYA = 2
PODPISI_DOCHITANNOGO = {
    "_nmck_poz": "расчёт НМЦК", "_nmck_tarif": "расчёт НМЦК", "_list_soglasovaniya": "лист согласования",
    "_nmck_metod": None, "_ploshchadka_podskazka": None, "_nomer_i_data_zayavki": None,
}


def _pusto(v) -> bool:
    return v is None or v == "" or v == [] or v == {}


def dochitat_iz_zayavki(nastrojki: Nastrojki, id_zakupki: str, soobshchit: bool = False) -> list[str]:
    """Заявка не менялась, но в карточке нет данных, которые программа теперь
    читает (расчёт НМЦК, источник финансирования, лист согласования…), —
    дочитать их, не трогая решения сотрудника. soobshchit — разовое
    сообщение «Дочитано: …» на карточке. Возвращает, что дочитано."""
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    if sostoyanie is None:
        raise ValueError(f"закупка {id_zakupki} не найдена")
    karta = sostoyanie.get("karta")
    put_zayavki = sostoyanie.get("put_zayavki")
    sostoyanie["versiya_chteniya"] = VERSIYA_CHTENIYA
    if not karta or not put_zayavki or not Path(put_zayavki).exists():
        _sohranit(nastrojki, id_zakupki, sostoyanie)
        return []
    try:
        svezhaya = _karta_v_slovar(_izvlech_iz_zayavki(Path(put_zayavki), nastrojki, None, sostoyanie.get("dokumenty")))
    except Exception:
        _sohranit(nastrojki, id_zakupki, sostoyanie)
        return []
    znach = karta.setdefault("znacheniya", {})
    istochniki = karta.setdefault("istochniki", {})
    dochitano: list[str] = []
    for imya, novoe in svezhaya["znacheniya"].items():
        if _pusto(novoe) or not _pusto(znach.get(imya)):
            continue
        if (istochniki.get(imya) or {}).get("fajl") == "sotrudnik":
            continue  # сотрудник очистил сам — не возвращать
        znach[imya] = novoe
        if imya in svezhaya.get("istochniki", {}):
            istochniki[imya] = svezhaya["istochniki"][imya]
        podpis = PODPISI_DOCHITANNOGO[imya] if imya in PODPISI_DOCHITANNOGO else (_podpis_polya(imya) or imya)
        if podpis:
            dochitano.append(podpis)
    stare, nove = karta.get("pozicii") or [], svezhaya.get("pozicii") or []
    if not stare and nove:
        karta["pozicii"] = nove
        dochitano.append("позиции")
    elif len(stare) == len(nove):
        stolbcy: list[str] = []
        for p_old, p_new in zip(stare, nove):
            for nazvanie, yach in p_new.items():
                old = p_old.get(nazvanie) or {}
                if old.get("fajl") == "sotrudnik" or str(old.get("znachenie") or "").strip():
                    continue
                if str(yach.get("znachenie") or "").strip():
                    p_old[nazvanie] = yach
                    stolbcy.append(nazvanie)
        for nazvanie in dict.fromkeys(stolbcy):
            dochitano.append(" ".join(nazvanie.split()).lower())
    # опечатки, которые сотрудник уже исправил («Исправить»), — и в дочитанном
    for najdeno, ispravlenie in (karta.get("ispravleniya_opechatok") or {}).items():
        _ispravit_slovo_vezde(karta, najdeno, ispravlenie)
    dochitano = list(dict.fromkeys(dochitano))
    if dochitano and soobshchit:
        sostoyanie["soobshchenie"] = "Дочитано из заявки: " + ", ".join(dochitano) + "."
    _sohranit(nastrojki, id_zakupki, sostoyanie)
    return dochitano


def _sopostavit_stolbec(staroe: str, vse: list[str], forma_426: bool) -> str | None:
    """Название столбца из сохранённого порядка (старой версии) → текущее
    название: то же, по списку замен, без учёта пробелов/регистра; столбец цены
    — по слову «цена»."""
    if staroe in vse:
        return staroe
    novoe = tablica_tz.pereimenovat_zagolovok(staroe, forma_426)
    if novoe in vse:
        return novoe
    norm = lambda s: " ".join(s.lower().split())  # noqa: E731
    for n in vse:
        if norm(n) in (norm(staroe), norm(novoe)):
            return n
    if "цена" in staroe.lower():
        return next((n for n in vse if "цена" in n.lower()), None)
    return None


def postroit_tablicu_tz(nastrojki: Nastrojki, sostoyanie: dict) -> tuple[list[str], list[list[str]], list[str], list[str]]:
    """Таблица ТЗ для документа — с учётом решений сотрудника на экране
    «Таблица ТЗ» (порядок, скрытые/возвращённые столбцы, добавленные вручную).
    Возвращает (видимые заголовки, строки, все доступные заголовки, скрытые)."""
    pozicii = (sostoyanie.get("karta") or {}).get("pozicii") or []
    znach = (sostoyanie.get("karta") or {}).get("znacheniya") or {}
    zagolovki, stroki = sobrat_tablicu_tz(pozicii)
    zagolovki, stroki, dobavleno = dobavit_otsutstvuyushchie_stolbcy(
        zagolovki, stroki,
        garantijnyj_srok=znach.get("_garantijnyj_srok_dlya_tablicy"),
        stavka_nds=znach.get("_stavka_nds_dlya_tablicy"),
    )
    znacheniya_po_stolbcu: dict[str, list[str]] = {
        n: [stroka[i] for stroka in stroki] for i, n in enumerate(zagolovki)
    }

    # служебные столбцы заявки (не идут в документ по умолчанию, но их можно
    # вернуть — раздел 5 ТЗ) — значения берём как есть, без переименования
    sluzhebnye_nazvaniya = _sluzhebnye_stolbcy_zayavki(pozicii)
    for n in sluzhebnye_nazvaniya:
        znacheniya_po_stolbcu[n] = [p[n]["znachenie"] for p in pozicii]

    # столбцы, добавленные сотрудником вручную («+ Добавить столбец»)
    dobavlennye_sotrudnikom: dict[str, list[str]] = sostoyanie.get("tz_dobavlennye") or {}
    for n, znacheniya in dobavlennye_sotrudnikom.items():
        znacheniya_po_stolbcu[n] = list(znacheniya) + [""] * (len(pozicii) - len(znacheniya))

    imeet_1875 = tablica_tz.NAZVANIE_STOLBCA_1875 in zagolovki
    zagolovki_bez_1875 = [z for z in zagolovki if z != tablica_tz.NAZVANIE_STOLBCA_1875]
    poryadok_po_umolchaniyu = zagolovki_bez_1875 + sluzhebnye_nazvaniya + list(dobavlennye_sotrudnikom.keys())
    if imeet_1875:
        poryadok_po_umolchaniyu.append(tablica_tz.NAZVANIE_STOLBCA_1875)

    vse_zagolovki = list(dict.fromkeys(poryadok_po_umolchaniyu))

    # порядок/скрытые столбцы, настроенные сотрудником (экран «Таблица ТЗ»);
    # пока сотрудник ничего не менял — скрыты служебные столбцы по умолчанию
    # названия из сохранённого порядка старой версии сопоставляются с текущими
    forma_426 = tablica_tz.eto_forma_426(list(pozicii[0].keys()) if pozicii else [])
    sopostavit = lambda spisok: [m for m in (_sopostavit_stolbec(n, vse_zagolovki, forma_426) for n in spisok) if m]  # noqa: E731
    poryadok = sostoyanie.get("tz_poryadok")
    if poryadok:
        poryadok_dejstvitelnyj = list(dict.fromkeys(sopostavit(poryadok)))
        # новые столбцы — на свои места по умолчанию (после предыдущего столбца программы)
        for i, n in enumerate(vse_zagolovki):
            if n not in poryadok_dejstvitelnyj:
                pered = next((vse_zagolovki[j] for j in range(i - 1, -1, -1) if vse_zagolovki[j] in poryadok_dejstvitelnyj), None)
                poryadok_dejstvitelnyj.insert(poryadok_dejstvitelnyj.index(pered) + 1 if pered else 0, n)
    else:
        poryadok_dejstvitelnyj = list(vse_zagolovki)
    if imeet_1875:  # мера по ПП № 1875 — всегда последним (раздел 5 ТЗ)
        poryadok_dejstvitelnyj = [n for n in poryadok_dejstvitelnyj if n != tablica_tz.NAZVANIE_STOLBCA_1875]
        poryadok_dejstvitelnyj.append(tablica_tz.NAZVANIE_STOLBCA_1875)

    # служебные столбцы (источник финансирования, ОКПД2, сумма…) не идут в
    # документ, пока сотрудник сам их не вернул
    sohranennye_skrytye = sostoyanie.get("tz_skrytye")
    vozvrashcheny = set(sopostavit(sostoyanie.get("tz_vozvrashcheny") or []))
    skrytye = set(sopostavit(sohranennye_skrytye)) if sohranennye_skrytye is not None else set()
    skrytye |= set(sluzhebnye_nazvaniya) - vozvrashcheny
    skrytye &= set(vse_zagolovki)

    vidimye_nazvaniya = [n for n in poryadok_dejstvitelnyj if n not in skrytye]
    novye_stroki = [[znacheniya_po_stolbcu[n][i] for n in vidimye_nazvaniya] for i in range(len(pozicii))]
    # все столбцы — в действующем порядке (экран «Таблица ТЗ» и перестановка)
    return vidimye_nazvaniya, novye_stroki, poryadok_dejstvitelnyj, list(skrytye)


def tz_izmenit(nastrojki: Nastrojki, id_zakupki: str, dejstvie: str, nazvanie: str, napravlenie: str | None = None) -> dict:
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    _, _, vse, skrytye_tekushchie = postroit_tablicu_tz(nastrojki, sostoyanie)
    poryadok = list(vse)  # действующий порядок, названия — текущие
    skrytye = set(skrytye_tekushchie)

    vozvrashcheny = set(sostoyanie.get("tz_vozvrashcheny") or [])
    if dejstvie == "ubrat":
        skrytye.add(nazvanie)
        vozvrashcheny.discard(nazvanie)
    elif dejstvie == "vernut":
        skrytye.discard(nazvanie)
        vozvrashcheny.add(nazvanie)
    elif dejstvie == "peremestit" and nazvanie in poryadok:
        i = poryadok.index(nazvanie)
        j = i - 1 if napravlenie == "vverh" else i + 1
        if 0 <= j < len(poryadok):
            poryadok[i], poryadok[j] = poryadok[j], poryadok[i]
    elif dejstvie == "sbros":
        # «Вернуть, как собрала программа» — убрать решения сотрудника по
        # порядку и видимости столбцов таблицы ТЗ
        poryadok = None
        skrytye = None  # пересчитается заново (служебные по умолчанию)
        sostoyanie["tz_dobavlennye"] = {}
        vozvrashcheny = set()
    elif dejstvie == "dobavit" and nazvanie.strip():
        pozicii = (sostoyanie.get("karta") or {}).get("pozicii") or []
        dobavlennye = sostoyanie.setdefault("tz_dobavlennye", {})
        if nazvanie not in dobavlennye:
            dobavlennye[nazvanie] = [""] * len(pozicii)
            poryadok.append(nazvanie)

    sostoyanie["tz_poryadok"] = poryadok
    sostoyanie["tz_skrytye"] = list(skrytye) if skrytye is not None else None
    sostoyanie["tz_vozvrashcheny"] = sorted(vozvrashcheny)
    _sohranit(nastrojki, id_zakupki, sostoyanie)
    return sostoyanie


def tz_izmenit_znachenie(nastrojki: Nastrojki, id_zakupki: str, nazvanie_stolbca: str, indeks_pozicii: int, znachenie: str) -> dict:
    """Значение ячейки добавленного вручную столбца таблицы ТЗ (построчно,
    только для столбцов, которых нет в заявке)."""
    sostoyanie = _zagruzit(nastrojki, id_zakupki)
    dobavlennye = sostoyanie.setdefault("tz_dobavlennye", {})
    if nazvanie_stolbca in dobavlennye and 0 <= indeks_pozicii < len(dobavlennye[nazvanie_stolbca]):
        dobavlennye[nazvanie_stolbca][indeks_pozicii] = znachenie
    _sohranit(nastrojki, id_zakupki, sostoyanie)
    return sostoyanie


# ------------------------------------------------------------------ даты, календарь

DNI_NEDELI_IMENITELNYJ = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"]


def data_razmeshcheniya_po_umolchaniyu(nastrojki: Nastrojki) -> tuple[date, bool]:
    """Дата размещения по умолчанию — сегодня; если сегодня нерабочий день —
    ближайший следующий рабочий (раздел 8 ТЗ, решение специалиста). Второй
    элемент — True, если дату перенесли."""
    kal = zagruzit_kalendar(nastrojki.sp("Производственный календарь.xlsx"))
    d = date.today()
    if kal.rabochij_den(d):
        return d, False
    perenesena = d
    while not kal.rabochij_den(perenesena):
        perenesena += timedelta(days=1)
    return perenesena, True


def pervaya_dostupnaya_data_podachi(nastrojki: Nastrojki, d_razmeshcheniya: date) -> date:
    """Первая дата окончания подачи, которую можно выбрать в календаре —
    15 полных дней после размещения, затем первый рабочий день (раздел 8 ТЗ)."""
    kal = zagruzit_kalendar(nastrojki.sp("Производственный календарь.xlsx"))
    pervyj_dostupnyj = d_razmeshcheniya + timedelta(days=16)
    while not kal.rabochij_den(pervyj_dostupnyj):
        pervyj_dostupnyj += timedelta(days=1)
    return pervyj_dostupnyj


def poyasnenie_sroka_podachi(nastrojki: Nastrojki, d_razmeshcheniya: date) -> str:
    """«Первая доступная дата — 26.10.2026, понедельник: 15 полных дней с
    10.10 по 24.10, 25.10 — воскресенье» — короткое пояснение справа от
    календаря (раздел 8 ТЗ)."""
    pervyj_dostupnyj = pervaya_dostupnaya_data_podachi(nastrojki, d_razmeshcheniya)
    nachalo_15_dnej = d_razmeshcheniya + timedelta(days=1)
    konec_15_dnej = d_razmeshcheniya + timedelta(days=15)
    den_posle = konec_15_dnej + timedelta(days=1)
    nedelya = DNI_NEDELI_IMENITELNYJ[pervyj_dostupnyj.weekday()]
    tekst = (
        f"Первая доступная дата — {pervyj_dostupnyj:%d.%m.%Y}, {nedelya}: "
        f"15 полных дней с {nachalo_15_dnej:%d.%m} по {konec_15_dnej:%d.%m}"
    )
    if den_posle != pervyj_dostupnyj:
        tekst += f", {den_posle:%d.%m} — {DNI_NEDELI_IMENITELNYJ[den_posle.weekday()]}"
    return tekst


def dni_mesyaca_dlya_podachi(nastrojki: Nastrojki, d_razmeshcheniya: date, god: int, mesyac: int, vybrannaya: date | None) -> list[dict]:
    """Список дней месяца с доступностью для даты окончания подачи заявок
    (раздел 8 ТЗ: 15 полных дней после размещения, затем рабочий день)."""
    kal = zagruzit_kalendar(nastrojki.sp("Производственный календарь.xlsx"))
    pervyj_dostupnyj = d_razmeshcheniya + timedelta(days=16)
    while not kal.rabochij_den(pervyj_dostupnyj):
        pervyj_dostupnyj += timedelta(days=1)

    _, dnej_v_mesyace = calendar_std.monthrange(god, mesyac)
    pervyj_den_nedeli, _ = calendar_std.monthrange(god, mesyac)
    rezultat = []
    for den in range(1, dnej_v_mesyace + 1):
        d = date(god, mesyac, den)
        dostupen = d >= pervyj_dostupnyj and kal.rabochij_den(d)
        if d < pervyj_dostupnyj:
            prichina = "не прошло 15 полных дней после размещения"
        elif not kal.rabochij_den(d):
            if d in kal.iskhlyucheniya:
                prichina = "праздник или перенесённый выходной (по производственному календарю)"
            elif d.weekday() == 5:
                prichina = "суббота"
            elif d.weekday() == 6:
                prichina = "воскресенье"
            else:
                prichina = "нерабочий день"
        else:
            prichina = ""
        rezultat.append({
            "den": den, "iso": d.isoformat(), "dostupen": dostupen, "prichina": prichina,
            "vybrana": vybrannaya == d,
        })
    return rezultat
