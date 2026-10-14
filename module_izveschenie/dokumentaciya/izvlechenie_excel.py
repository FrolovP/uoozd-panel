# -*- coding: utf-8 -*-
"""Извлечение данных карточки закупки из заявки в формате Excel (раздел 9 ТЗ).

Значения ищутся по подписям в заявке (текст меток, заголовки таблиц), а не по
адресам ячеек: у разных подразделений одни и те же сведения стоят в разных
ячейках. Для каждого найденного значения запоминается источник — файл, лист,
ячейка и дословная цитата, — это видно сотруднику в карточке (раздел 4 ТЗ).

Модуль ядра: не знает ни про программу, ни про браузер (раздел 14 ТЗ).
Получает путь к файлу заявки и путь к справочникам, отдаёт обычные словари.
Путей внутри модуля не хранится ни одного — все приходят параметрами.

Поля, которые модуль НЕ заполняет на этом шаге (выбор сотрудника, справочник
подписантов, даты, переключатели нацрежима и требований по умолчанию —
разделы 6, 7, 8 ТЗ): площадка (только подсказка), все даты, номер закупки,
контакты ответственного по п. 1, год, nr_*, dok_1875_*, ru_minzdrav,
девять требований treb_* по умолчанию, суммы обеспечения исполнения договора.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

import pravila


@dataclass
class Istochnik:
    """Источник одного значения карточки: где в заявке оно нашлось."""

    fajl: str
    list_: str
    yacheyka: str
    citata: str


@dataclass
class Kartochka:
    """Результат извлечения: данные для шаблона + источник каждого поля."""

    znacheniya: dict[str, Any] = field(default_factory=dict)
    istochniki: dict[str, Istochnik] = field(default_factory=dict)
    pozicii: list[dict[str, Any]] = field(default_factory=list)
    preduprezhdeniya: list[str] = field(default_factory=list)


def _text(cell: Cell | None) -> str:
    return "" if cell is None or cell.value is None else str(cell.value).strip()


def naiti_podpis(ws: Worksheet, *fragmenty: str) -> Cell | None:
    """Первая ячейка, чей текст содержит все фрагменты (без учёта регистра)."""
    frag_low = [f.lower() for f in fragmenty]
    for row in ws.iter_rows():
        for c in row:
            v = _text(c).lower()
            if v and all(f in v for f in frag_low):
                return c
    return None


def naiti_po_prefiksu(ws: Worksheet, prefiks: str) -> Cell | None:
    """Первая ячейка, чей текст начинается с указанного префикса."""
    prefiks_low = prefiks.lower()
    for row in ws.iter_rows():
        for c in row:
            if _text(c).lower().startswith(prefiks_low):
                return c
    return None


def naiti_tochnoe_sovpadenie(ws: Worksheet, tekst: str) -> Cell | None:
    """Первая ячейка, чей текст в точности равен указанному (без учёта пробелов)."""
    for row in ws.iter_rows():
        for c in row:
            if _text(c) == tekst:
                return c
    return None


def sleduyuschaya_znachimaya_yacheyka(ws: Worksheet, posle: Cell | None) -> Cell | None:
    """Следующая непустая ячейка после данной, в порядке чтения документа
    (слева направо, затем сверху вниз).

    Так в заявке устроены пары «подпись — значение»: значение стоит сразу
    после подписи — в её строке правее или на следующей непустой строке.
    Поэтому значения ищутся по подписям, а не по адресам ячеек (раздел 9 ТЗ).
    """
    if posle is None:
        return None
    for row in ws.iter_rows(min_row=posle.row):
        for c in row:
            if c.row == posle.row and c.column <= posle.column:
                continue
            if _text(c):
                return c
    return None


def znachenie_po_podpisi(ws: Worksheet, *podpisi: str) -> tuple[Cell | None, str]:
    """Значение по первой найденной подписи из вариантов (формы заявок
    различаются). Если в ячейке подписи после неё есть текст («Место
    доставки: г. Москва…») — значение в той же ячейке, иначе — следующая
    непустая ячейка (подпись и значение в разных ячейках)."""
    for podpis in podpisi:
        label = naiti_podpis(ws, podpis)
        if label is None:
            continue
        tekst = _text(label)
        nachalo = tekst.lower().find(podpis.lower())
        ostatok = tekst[nachalo + len(podpis):].lstrip(" :\n\t")
        if ostatok:
            return label, ostatok
        cell = sleduyuschaya_znachimaya_yacheyka(ws, label)
        return cell, _text(cell)
    return None, ""


def sobrat_do(ws: Worksheet, posle: Cell | None, stop_prefiks: str) -> list[str]:
    """Собрать текст всех непустых ячеек после данной, пока следующая не
    начинается с указанного префикса (её саму не включать)."""
    rezultat: list[str] = []
    tekushchaya = posle
    while True:
        sleduyuschaya = sleduyuschaya_znachimaya_yacheyka(ws, tekushchaya)
        if sleduyuschaya is None:
            break
        tekst = _text(sleduyuschaya)
        if tekst.lower().startswith(stop_prefiks.lower()):
            break
        rezultat.append(tekst)
        tekushchaya = sleduyuschaya
    return rezultat


def sobrat_poka_pohozhe(ws: Worksheet, posle: Cell | None, prefiksy: tuple[str, ...]) -> list[str]:
    """Собрать текст непустых ячеек после данной, пока очередная начинается
    с одного из указанных префиксов (так несколько ячеек подряд образуют один
    смысловой блок, например про нацрежим)."""
    rezultat: list[str] = []
    tekushchaya = posle
    while True:
        sleduyuschaya = sleduyuschaya_znachimaya_yacheyka(ws, tekushchaya)
        if sleduyuschaya is None:
            break
        tekst = _text(sleduyuschaya)
        if not any(tekst.lower().startswith(p.lower()) for p in prefiksy):
            break
        rezultat.append(tekst)
        tekushchaya = sleduyuschaya
    return rezultat


_CHISLOVAYA_STROKA = re.compile(r"^-?\d[\d ]*(?:[.,]\d+)?$")


def format_chislo(znachenie: Any) -> str:
    """Число — в вид «13 712,07» с неразрывным пробелом между разрядами
    (раздел 5 ТЗ). Если в заявке число уже напечатано строкой с обычным
    пробелом, пробел заменяется на неразрывный; остальной текст не трогается."""
    if isinstance(znachenie, str):
        tekst = znachenie.strip()
        if _CHISLOVAYA_STROKA.match(tekst):
            return tekst.replace(" ", " ")
        return tekst
    d = Decimal(str(znachenie)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    znak = "-" if d < 0 else ""
    d = abs(d)
    celaya, kopejki = f"{d:.2f}".split(".")
    gruppy = []
    while len(celaya) > 3:
        gruppy.insert(0, celaya[-3:])
        celaya = celaya[:-3]
    gruppy.insert(0, celaya)
    return znak + " ".join(gruppy) + "," + kopejki


def s_nerazryvnymi_probelami_v_chislah(tekst: str) -> str:
    """В строке с текстом вокруг числа (например, «в т.ч. НДС 10% - 161 182,46
    руб.») заменяет обычный пробел между разрядами числа на неразрывный
    (раздел 5 ТЗ); остальные пробелы не трогает."""
    return re.sub(r"(?<=\d) (?=\d)", " ", tekst)


def zagruzit_podpisantov(put: str | Path) -> dict[str, tuple[str, str]]:
    """Справочник подписантов титульного листа — Подписанты.xlsx, лист
    «Подписанты» (столбцы: Должность, И.О. Фамилия, По умолчанию утверждает,
    Роль — «утверждает»/«согласует»). Утверждающий по умолчанию — строка с
    отметкой «да»; если он отсутствует, сотрудник выбирает другого из списка.
    Все справочники — только Excel, лист «Как заполнять» программа не читает
    (раздел 4 ТЗ)."""
    wb = openpyxl.load_workbook(put, data_only=True)
    ws = wb["Подписанты"]
    rezultat: dict[str, tuple[str, str]] = {}
    vse: list[tuple[str, str]] = []
    vse_s_rolyu: list[dict[str, str]] = []
    for dolzhnost, fio, po_umolchaniyu, rol in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        if not dolzhnost:
            continue
        vse.append((str(dolzhnost).strip(), str(fio).strip()))
        vse_s_rolyu.append({"dolzhnost": str(dolzhnost).strip(), "fio": str(fio).strip(), "rol": str(rol or "").strip().lower()})
        if str(po_umolchaniyu or "").strip().lower() == "да":
            rezultat["utverzhdayuschij_po_umolchaniyu"] = (str(dolzhnost).strip(), str(fio).strip())
    rezultat["vse"] = vse
    rezultat["vse_s_rolyu"] = vse_s_rolyu
    return rezultat


def zagruzit_podrazdeleniya(put: str | Path) -> dict[str, tuple[str, str]]:
    """Справочник «подразделение -> первый согласующий» — «Подразделение —
    согласующий.xlsx», лист «Подразделения» (столбцы: сокращение в заявке,
    полное название в именительном падеже, полное название подразделения
    (родительный падеж), должность согласующего, И.О. Фамилия).

    В разных заявках подразделение называется по-разному («Подразделение КЦ»
    в Excel-форме, «Наименование подразделения: Клинический центр» в
    Word-форме) — поэтому ключей в итоговом словаре на одну запись может быть
    два: и сокращение, и полное название."""
    wb = openpyxl.load_workbook(put, data_only=True)
    ws = wb["Подразделения"]
    rezultat: dict[str, tuple[str, str]] = {}
    for sokr, nominativ, polnoe, dolzhnost, fio in ws.iter_rows(min_row=2, max_col=5, values_only=True):
        if not sokr:
            continue
        zapis = (f"{str(dolzhnost).strip()} {str(polnoe).strip()}", str(fio).strip())
        rezultat[str(sokr).strip()] = zapis
        if nominativ:
            rezultat[str(nominativ).strip()] = zapis
    return rezultat


def naiti_zapis_podrazdeleniya(tekst: str, podrazdeleniya: dict[str, tuple[str, str]]) -> tuple[str, str] | None:
    """Ищет в тексте заявки любой из ключей справочника подразделений
    (сокращение или полное название) — а не наоборот, чтобы не зависеть
    от того, как конкретная заявка называет своё подразделение."""
    nizhnij = tekst.lower()
    for kod, zapis in podrazdeleniya.items():
        if kod.lower() in nizhnij:
            return zapis
    return None


def zagruzit_sotrudnikov(put: str | Path) -> dict[str, str]:
    """Справочник сотрудников отдела — «Сотрудники отдела.xlsx», лист
    «Сотрудники» (столбцы: ФИО, Добавочный, Городской телефон, Телефон в
    извещении (п. 1)). Если телефон в извещении не заполнен, собирается из
    городского номера и добавочного: «8» + городской + «, доб. » + добавочный
    (добавочный печатается всегда — решение специалиста от 09.10.2026)."""
    wb = openpyxl.load_workbook(put, data_only=True)
    ws = wb["Сотрудники"]
    rezultat: dict[str, str] = {}
    for fio, dobavochnyj, gorodskoj, telefon_v_izveshchenii in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        if not fio:
            continue
        if telefon_v_izveshchenii:
            telefon = str(telefon_v_izveshchenii).strip()
        else:
            telefon = f"8{gorodskoj}, доб. {dobavochnyj}"
        rezultat[str(fio).strip()] = telefon
    return rezultat


def zagruzit_rukovoditelej(put: str | Path) -> set[str]:
    """ФИО руководителей из «Сотрудники отдела.xlsx»: столбец «Руководитель»
    = «да». Руководитель видит закупки всех сотрудников. Столбца нет — нет и
    руководителей."""
    wb = openpyxl.load_workbook(put, data_only=True)
    ws = wb["Сотрудники"]
    zagolovki = [str(c.value or "").strip().lower() for c in ws[1]]
    if "руководитель" not in zagolovki:
        return set()
    i = zagolovki.index("руководитель")
    return {
        str(stroka[0]).strip() for stroka in ws.iter_rows(min_row=2, values_only=True)
        if stroka and stroka[0] and len(stroka) > i and str(stroka[i] or "").strip().lower() == "да"
    }


def zagruzit_trebovaniya_standartnye(put: str | Path) -> set[str]:
    """«Требования к участникам (стандартные).xlsx» — первый столбец листа
    с данными (не «Как заполнять»), без заголовка."""
    wb = openpyxl.load_workbook(put, data_only=True)
    ws = wb.worksheets[0]
    return {
        str(row[0]).strip() for row in ws.iter_rows(min_row=2, max_col=1, values_only=True)
        if row[0]
    }


# заявка разделяет «а) ...», «б) ...» оставшиеся пункты фразы «Цена должна
# включать»; то, что уже напечатано в п. 8 шаблона, из фразы убирается
# (решение специалиста, раздел CLAUDE.md «Решения специалиста»)
_UZHE_V_SHABLONE_P8 = ("ндс", "обязательные платежи", "и др", "и пр")
_BUKVY_PEREЧNYA = "абвгдежзиклмнопрстуфхцчшщ"


def razobrat_cenu_vklyuchaet(tekst: str) -> list[str]:
    chasti = [ch.strip() for ch in tekst.rstrip(". ").split(";")]
    punkty = []
    for chast in chasti:
        slova_perechnya = [s.strip() for s in chast.split(",")]
        ostavshiesya = [
            s for s in slova_perechnya
            if s and not any(marker in s.lower() for marker in _UZHE_V_SHABLONE_P8)
        ]
        if ostavshiesya:
            punkty.append(", ".join(ostavshiesya))
    return [f"{_BUKVY_PEREЧNYA[i]}) {punkt}." for i, punkt in enumerate(punkty)]


class ZayavkaExcel:
    """Заявка в файле Excel (данные читаются с первого листа)."""

    def __init__(self, put: str | Path):
        self.put = Path(put)
        self.fajl = self.put.name
        wb = openpyxl.load_workbook(self.put, data_only=True)
        self.wb = wb
        self.ws = wb.worksheets[0]
        self.list_ = self.ws.title

    def _polozhit(self, karta: Kartochka, imya: str, cell: Cell | None, znachenie: Any, citata: str | None = None) -> None:
        karta.znacheniya[imya] = znachenie
        karta.istochniki[imya] = Istochnik(self.fajl, self.list_, cell.coordinate if cell is not None else "", citata or (_text(cell) if cell is not None else ""))

    # ------------------------------------------------------------ предмет закупки
    def _predmet(self, karta: Kartochka) -> None:
        # разные формы заявки: «Заявка на закрытый аукцион на …» или
        # «Закрытый аукцион в электронной форме на …»
        for prefiks in ("Заявка на закрытый аукцион на ", "Закрытый аукцион в электронной форме на "):
            cell = naiti_po_prefiksu(self.ws, prefiks)
            if cell is not None:
                break
        if cell is None:
            return
        tekst = _text(cell)
        vin = tekst[len(prefiks):].strip()
        pervoe_slovo_im = {"поставку": "Поставка"}
        pervoe, _, ostatok = vin.partition(" ")
        im = pervoe_slovo_im.get(pervoe.lower(), pervoe.capitalize()) + " " + ostatok
        self._polozhit(karta, "predmet_vin", cell, vin, tekst)
        self._polozhit(karta, "predmet_im", cell, im, tekst)

        nmcd_cell = sleduyuschaya_znachimaya_yacheyka(self.ws, cell)
        if nmcd_cell is not None and nmcd_cell.row == cell.row:
            self._polozhit(karta, "nmcd", nmcd_cell, format_chislo(nmcd_cell.value))

    # ------------------------------------------------------- площадка (подсказка)
    def _ploshchadka_podskazka(self, karta: Kartochka) -> None:
        cell = naiti_po_prefiksu(self.ws, "Электронная площадка")
        if cell is None:
            return
        tekst = _text(cell)[len("Электронная площадка"):].lstrip(" :").strip()
        if tekst:
            self._polozhit(karta, "_ploshchadka_podskazka", cell, tekst)

    # --------------------------------------------------- номер и дата заявки
    def _rekvizity_zayavki(self, karta: Kartochka) -> None:
        label = naiti_tochnoe_sovpadenie(self.ws, "Заявка")
        cell = sleduyuschaya_znachimaya_yacheyka(self.ws, label) if label else None
        if cell is not None:
            self._polozhit(karta, "_nomer_i_data_zayavki", cell, _text(cell))

    # -------------------------------------------------------------------- НДС
    def _nds_tekst(self, karta: Kartochka) -> None:
        cell = naiti_po_prefiksu(self.ws, "ИТОГО:")
        if cell is None:
            return  # другая форма заявки — _nds_iz_itogo_rub после позиций
        tekst = _text(cell)
        if "НДС не облагается" in tekst:
            znachenie = tekst[tekst.index("НДС не облагается"):]
        elif "в т.ч. НДС" in tekst:
            nachalo = tekst.index("в т.ч. НДС")
            konec = tekst.find(" от суммы", nachalo)
            znachenie = tekst[nachalo: konec if konec != -1 else len(tekst)]
        else:
            return
        znachenie = s_nerazryvnymi_probelami_v_chislah(znachenie.strip())
        self._polozhit(karta, "nds_tekst", cell, znachenie, tekst)

    # ---------------------------------------------------------- цена включает
    def _cena_vklyuchaet(self, karta: Kartochka) -> None:
        cell, tekst = znachenie_po_podpisi(self.ws, "Цена должна включать", "Цена товара включает")
        if cell is None or not tekst:
            return
        self._polozhit(karta, "cena_vklyuchaet", cell, razobrat_cenu_vklyuchaet(tekst), tekst)

    # ------------------------------------------------ контактное лицо (п. 1)
    def _kontakt(self, karta: Kartochka, sotrudniki: dict[str, str], polzovatel_fio: str | None) -> None:
        # ответственное лицо Заказчика — сотрудник, выбравший себя при первом
        # запуске программы (раздел 4 ТЗ), а не из заявки
        if not polzovatel_fio:
            return
        karta.znacheniya["kontakt_fio"] = polzovatel_fio
        karta.istochniki["kontakt_fio"] = Istochnik("spravochnik", "polzovatel", "", "выбор сотрудника при первом запуске")
        telefon = sotrudniki.get(polzovatel_fio)
        if telefon:
            karta.znacheniya["kontakt_tel"] = telefon
            karta.istochniki["kontakt_tel"] = Istochnik("spravochnik", "Сотрудники отдела", "", f"справочник сотрудников: {polzovatel_fio}")

    # ------------------------------------------------- подразделение заявки
    def _tekst_podrazdeleniya(self) -> str | None:
        cell = naiti_podpis(self.ws, "Подразделение")
        return _text(cell) if cell is not None else None

    # ------------------------------------------------- утверждающий, согласующие
    def _utverzhdenie(self, karta: Kartochka, podpisanty: dict[str, tuple[str, str]]) -> None:
        # утверждает всегда проректор по экономике и финансам из справочника
        # подписантов; заявка здесь не читается (решение специалиста от 09.10.2026)
        zapis = podpisanty.get("utverzhdayuschij_po_umolchaniyu")
        if zapis is None:
            return
        dolzhnost, fio = zapis
        karta.znacheniya["utverzhdayu"] = {"dolzhnost": dolzhnost, "fio": fio}
        karta.istochniki["utverzhdayu"] = Istochnik("spravochnik", "podpisanty", "", "справочник подписантов, по умолчанию")

    def _soglasovanie(self, karta: Kartochka, podrazdeleniya: dict[str, tuple[str, str]]) -> None:
        # лист согласования заявки — только для проверки: если выбранного
        # согласующего там нет, самопроверка предупреждает, но не меняет
        # значение (решение специалиста от 09.10.2026)
        label = naiti_podpis(self.ws, "Согласовано:") or naiti_podpis(self.ws, "Лист согласования")
        if label is not None:
            stroki = [label_tekst for label_tekst in sobrat_do(self.ws, label, "Исполнитель:") if label_tekst]
            self._polozhit(karta, "_list_soglasovaniya", label, stroki)

        # первый согласующий — по подразделению заявки, из справочника
        # «подразделение -> согласующий», а не из листа согласования заявки
        # (решение специалиста от 09.10.2026); остальных добавляет сотрудник
        tekst_podrazdeleniya = self._tekst_podrazdeleniya()
        zapis = naiti_zapis_podrazdeleniya(tekst_podrazdeleniya, podrazdeleniya) if tekst_podrazdeleniya else None
        if zapis is None:
            return
        dolzhnost, fio = zapis
        karta.znacheniya["soglasovano"] = [{"dolzhnost": dolzhnost, "fio": fio}]
        karta.istochniki["soglasovano"] = Istochnik("spravochnik", "podrazdelenie_soglasuyushchij", "", f"подразделение: {tekst_podrazdeleniya}")

    # ------------------------------------------------- условия для проекта договора
    def _usloviya_dogovora(self, karta: Kartochka) -> None:
        # раздел 9 ТЗ: в саму документацию не идут, сохраняются в карточке для договора
        for podpisi, imya in (
            (("Место поставки:", "Место доставки:"), "mesto_postavki"),
            (("Сроки (график) поставки:", "Условия поставки:"), "sroki_postavki"),
            (("Сроки и условия оплаты:",), "usloviya_oplaty"),
        ):
            cell, tekst = znachenie_po_podpisi(self.ws, *podpisi)
            if cell is not None and tekst:
                self._polozhit(karta, imya, cell, tekst, _text(cell))

        cell = naiti_po_prefiksu(self.ws, "Срок действия Договора:")
        if cell is not None:
            tekst = _text(cell)
            self._polozhit(karta, "srok_dejstviya_dogovora", cell, tekst.split(":", 1)[1].strip(), tekst)

    # ------------------------------------------------- ответственный за договор
    def _otv_dogovor(self, karta: Kartochka) -> None:
        label = naiti_podpis(self.ws, "Ответственный исполнитель по договору")
        cell = sleduyuschaya_znachimaya_yacheyka(self.ws, label) if label else None
        if cell is not None:
            self._polozhit(karta, "otv_dogovor", cell, _text(cell))
        else:
            self._otv_dogovor_iz_obyazatelstva(karta)

        # телефон стоит отдельной строкой ниже подписи — своя подпись «тел.»
        tel_cell = naiti_po_prefiksu(self.ws, "тел.")
        if tel_cell is not None:
            tekst = _text(tel_cell)
            self._polozhit(karta, "otv_dogovor_telefon", tel_cell, tekst[len("тел."):].strip(), tekst)

    def _otv_dogovor_iz_obyazatelstva(self, karta: Kartochka) -> None:
        """Другая форма заявки: «Я, Марков А.Г., являясь ответственным
        исполнителем по Договору…» — должность берётся из строки листа
        согласования с этой фамилией: «Доцент … Марков А.Г.»."""
        cell = naiti_podpis(self.ws, "являясь ответственным исполнителем")
        m = re.search(r"Я,\s*([^,]+?),\s*являясь ответственным исполнителем", _text(cell)) if cell else None
        if not m:
            return
        fio = m.group(1).strip()
        familiya = fio.split()[0]
        list_soglasovaniya = naiti_podpis(self.ws, "Лист согласования")
        nachalo = list_soglasovaniya.row if list_soglasovaniya else 1
        for row in self.ws.iter_rows(min_row=nachalo, max_row=cell.row - 1):
            yach = [c for c in row if _text(c)]
            if len(yach) >= 2 and familiya in _text(yach[-1]):
                # должность — самая длинная ячейка строки без фамилии (номер строки не в счёт)
                kandidaty = [c for c in yach if familiya not in _text(c) and not _text(c).replace(".", "").isdigit()]
                if not kandidaty:
                    continue
                dolzh_cell = max(kandidaty, key=lambda c: len(_text(c)))
                dolzhnost = " ".join(_text(dolzh_cell).split())
                znachenie = f"{dolzhnost[:1].upper()}{dolzhnost[1:]} {fio}"
                self._polozhit(karta, "otv_dogovor", dolzh_cell, znachenie, f"{_text(cell)[:120]}… / {dolzhnost} {_text(yach[-1])}")
                return
        self._polozhit(karta, "otv_dogovor", cell, fio, _text(cell)[:200])

    # ------------------------------------------------------------- инструкция
    def _instrukciya(self, karta: Kartochka) -> None:
        label = naiti_podpis(self.ws, "Инструкция по описанию товара")
        cell = sleduyuschaya_znachimaya_yacheyka(self.ws, label) if label else None
        if cell is None:
            return
        abzacy = [a.strip() for a in _text(cell).split("\n") if a.strip()]
        self._polozhit(karta, "instrukciya_tovara", cell, abzacy, _text(cell))

    # ------------------------------------------------------------ обеспечение
    def _obespechenie(self, karta: Kartochka) -> None:
        cell, tekst = znachenie_po_podpisi(self.ws, "Обеспечение заявки и договора", "Обеспечение заявки, договора:")
        if cell is None or not tekst:
            return
        self._polozhit(karta, "obespechenie_tekst", cell, tekst)

        zayavka_ne = re.search(r"обеспечение заявки\s*-\s*не (применяется|требуется)", tekst, re.I)
        self._polozhit(karta, "obesp_zayavki", cell, {"ustanovleno": not bool(zayavka_ne)}, tekst)

        dogovor_procent = re.search(r"обеспечение договора\s*-\s*(\d+(?:[.,]\d+)?)\s*%", tekst, re.I)
        dogovor_ne = re.search(r"обеспечение договора\s*-\s*не (применяется|требуется)", tekst, re.I)
        if dogovor_procent:
            self._polozhit(karta, "obesp_isp_procent", cell, dogovor_procent.group(1).replace(".", ","), tekst)
        elif dogovor_ne:
            self._polozhit(karta, "obesp_isp_procent", cell, "", tekst)

    # ------------------------------------------------------------- текст о нацрежиме
    def _nacrezhim(self, karta: Kartochka) -> None:
        label = naiti_podpis(self.ws, "Иные условия исполнения договора")
        if label is None:
            return
        prefiksy = ("Постановление", "ПОСТАНОВЛЕНИЕ", "Наименование товара:", "Установлен вид требования:")
        stroki = sobrat_poka_pohozhe(self.ws, label, prefiksy)
        if stroki:
            posledstvennaya = sleduyuschaya_znachimaya_yacheyka(self.ws, label)
            self._polozhit(karta, "nacrezhim_tekst", posledstvennaya, " ".join(stroki), " ".join(stroki))

    # --------------------------------------------------- требования к участникам
    def _trebovaniya_k_uchastnikam(self, karta: Kartochka, standartnye: set[str]) -> None:
        label = naiti_podpis(self.ws, "Требования к участникам:") or naiti_po_prefiksu(self.ws, "Требования к участникам в соответствии")
        if label is None:
            return
        stroki = sobrat_do(self.ws, label, "Начальная максимальная цена")
        self._polozhit(karta, "trebovaniya_k_uchastnikam", label, stroki, "\n".join(stroki))

        # строка вида «Регистрационные удостоверения — не требуются» — это
        # отсутствие требования, а не дополнительное требование
        dopolnitelnye = [
            s for s in stroki
            if not pravila.sovpadaet_so_standartnym(s, standartnye) and not re.search(r"\bне\s+требу", s, re.I)
        ]
        if dopolnitelnye:
            tekst = dopolnitelnye[0].rstrip(".;")
            if not tekst.endswith(";"):
                tekst += ";"
            self._polozhit(karta, "dok_sootv_uchastnika", label, "требуется")
            self._polozhit(karta, "dok_sootv_tekst", label, tekst, dopolnitelnye[0])
        else:
            self._polozhit(karta, "dok_sootv_uchastnika", label, "не требуется")
            self._polozhit(karta, "dok_sootv_tekst", label, "")

    # ------------------------------------------------------------- расчёт НМЦК
    def _nmck_raschet(self, karta: Kartochka) -> None:
        if naiti_podpis(self.ws, "Используемый метод определения НМЦК") is None:
            # в другой форме заявки расчёт НМЦК — на отдельном листе («Расчет НМЦК»)
            for ws in self.wb.worksheets[1:]:
                if naiti_podpis(ws, "Используемый метод определения НМЦК") is not None:
                    ws_osnovnoj, list_osnovnoj = self.ws, self.list_
                    self.ws, self.list_ = ws, ws.title
                    try:
                        self._nmck_raschet(karta)
                    finally:
                        self.ws, self.list_ = ws_osnovnoj, list_osnovnoj
                    return
        label = naiti_podpis(self.ws, "Используемый метод определения НМЦК")
        if label is None:
            return
        metod_cell = sleduyuschaya_znachimaya_yacheyka(self.ws, label)
        metod = _text(metod_cell).lower()

        raschet: dict[str, Any] = {
            "smeta": "смет" in metod,
            "po_minimalnomu": "минимальн" in metod,
        }

        if any(slovo in metod for slovo in ("метод", "тариф", "смет")):
            karta.znacheniya["_nmck_metod"] = _text(metod_cell)
        if "тарифный" in metod:
            # тарифный метод (лекарства по Госреестру): строки расчёта — в _nmck_tarif
            raschet["po_obshchej_summe"] = False
            self._polozhit(karta, "nmck_raschet", metod_cell, raschet, _text(metod_cell))
            return

        predlozheniya = []
        obrazec = re.compile(r"Предложение\s+\d+\s+от\s+(\d{2}\.\d{2}\.\d{4})\s*№\s*(\S+)", re.I)
        for row in self.ws.iter_rows():
            for c in row:
                m = obrazec.match(_text(c))
                if not m:
                    continue
                summa_cell = sleduyuschaya_znachimaya_yacheyka(self.ws, c)
                predlozheniya.append({
                    "vh_data": f"{m.group(1)}г.",
                    "vh_nomer": m.group(2),
                    "summa": format_chislo(summa_cell.value) if summa_cell is not None else "",
                })

        raschet["po_obshchej_summe"] = bool(predlozheniya)
        if predlozheniya:
            raschet["kp"] = predlozheniya
            obekt_label = naiti_podpis(self.ws, "Основные характеристики объекта закупки")
            if obekt_label is not None:
                obekt_cell = sleduyuschaya_znachimaya_yacheyka(self.ws, obekt_label)
                raschet["obekt"] = _text(obekt_cell)
            for podpis, imya in (
                ("Средняя арифметическая величина", "srednyaya"),
                ("Количество значений", "chislo_znachenij"),
                ("Количество (объем) ТРУ", "obem"),
                ("Коэф.вариации", "koef_var"),
                ("Совокупность значений", "odnorodnost"),
            ):
                l = naiti_podpis(self.ws, podpis)
                if l is None:
                    continue
                c = sleduyuschaya_znachimaya_yacheyka(self.ws, l)
                if c is not None:
                    raschet[imya] = format_chislo(c.value)
            sigma_label = naiti_po_prefiksu(self.ws, "σ")
            if sigma_label is not None:
                c = sleduyuschaya_znachimaya_yacheyka(self.ws, sigma_label)
                if c is not None:
                    raschet["sko"] = format_chislo(c.value)

        self._polozhit(karta, "nmck_raschet", metod_cell, raschet, _text(metod_cell))

    # ------------------------------ расчёт НМЦК по позициям (таблица листа)
    def _nmck_po_poziciyam(self, karta: Kartochka) -> None:
        """Таблица расчёта НМЦК «по позициям» (форма с листом «Расчет НМЦК»,
        как 426зэа): поставщики с входящими номерами КП, цены по позициям,
        n, v, средняя, V, σ, НМЦК по позиции и ИТОГО — вставляется в
        Приложение № 2 как есть, числа — «2 952,33»."""
        raschet = karta.znacheniya.get("nmck_raschet") or {}
        if raschet.get("smeta") or raschet.get("po_obshchej_summe") or "_nmck_tarif" in karta.znacheniya:
            return
        for ws in self.wb.worksheets:
            hdr = naiti_podpis(ws, "Средняя цена за единицу") or naiti_podpis(ws, "НМЦКрын")
            if hdr is not None:
                break
        else:
            return
        shapka = [c for c in ws[hdr.row] if _text(c)]
        pervyj, poslednij = shapka[0].column, max(c.column for c in shapka)
        stolbcy = list(range(pervyj, poslednij + 1))

        def tekst(r: int, c: int) -> str:
            return _text(ws.cell(r, c))

        verh = [tekst(hdr.row, c) for c in stolbcy]
        niz = [tekst(hdr.row + 1, c) for c in stolbcy]
        est_podshapka = not niz[0].replace(".0", "").isdigit() and any(niz)
        # опечатка в годе даты КП («11.06.20266») — исправляется, сотрудник видит предупреждение
        ispravleny_daty = []
        if est_podshapka:
            for i, s in enumerate(niz):
                m = re.search(r"(?<!\d)(\d{1,2}\.\d{2}\.)(20\d{2})\d+(?!\d)", s)
                if m:
                    niz[i] = s[:m.start()] + m.group(1) + m.group(2) + s[m.end():]
                    ispravleny_daty.append({"bylo": m.group(0), "stalo": m.group(1) + m.group(2),
                                            "yacheyka": ws.cell(hdr.row + 1, stolbcy[i]).coordinate, "list": ws.title})
        stroki: list[list[str]] = []
        itogo, itogo_stolbec = None, None
        nachalo = hdr.row + (2 if est_podshapka else 1)
        for r in range(nachalo, nachalo + 200):
            znach = [ws.cell(r, c).value for c in stolbcy]
            nomer = str(znach[0] if znach[0] is not None else "").strip()
            if re.fullmatch(r"\d+(\.0)?", nomer):
                stroka = []
                for i, v in enumerate(znach):
                    if i == 0:
                        stroka.append(nomer.replace(".0", ""))
                    elif isinstance(v, (int, float)) and "количество" in verh[i].lower():
                        stroka.append(format_chislo(v).removesuffix(",00"))
                    elif isinstance(v, (int, float)):
                        stroka.append(format_chislo(v))
                    else:
                        stroka.append("" if v is None else str(v).strip())
                stroki.append(stroka)
            elif stroki:
                chisla = [(i, v) for i, v in enumerate(znach) if isinstance(v, (int, float))]
                if chisla:  # строка «ИТОГО:» (пустые строки перед ней пропускаются)
                    itogo_stolbec, itogo = chisla[-1][0], format_chislo(chisla[-1][1])
                    break
                if any(v not in (None, "") for v in znach):
                    break
        if not stroki:
            return
        metod = next((s for s in verh + niz if "метод" in s.lower()), "")
        if metod:
            karta.znacheniya["_nmck_metod"] = metod
        self.list_, list_osnovnoj = ws.title, self.list_
        try:
            self._polozhit(karta, "_nmck_poz", hdr, {
                "verh": verh, "niz": niz if est_podshapka else None,
                "stroki": stroki, "itogo": itogo, "itogo_stolbec": itogo_stolbec,
                "ispravleny_daty": ispravleny_daty,
            })
        finally:
            self.list_ = list_osnovnoj

    # ------------------------------------------- тарифный расчёт НМЦК (строки)
    def _nmck_tarif(self, karta: Kartochka) -> None:
        """Строки тарифного расчёта (лекарства по Госреестру) — как в
        shablon/skripty_dlya_spravki/nmck_tarif.py. В kartochka_ozhidaemaya.json
        не сравниваются: расчёт вставляется в Приложение № 2 отдельной таблицей
        (вложенный документ), не текстовыми полями."""
        label = naiti_podpis(self.ws, "Используемый метод определения НМЦК")
        if label is None:
            return
        metod_cell = sleduyuschaya_znachimaya_yacheyka(self.ws, label)
        if metod_cell is None or "тарифный" not in _text(metod_cell).lower():
            return

        hdr = naiti_podpis(self.ws, "Расчет НМЦК")
        if hdr is None:
            return
        stolbcy = [(c.column, _text(c)) for c in self.ws[hdr.row] if _text(c)]
        if len(stolbcy) < 6:
            return
        col_ids = [c for c, _ in stolbcy]

        pozicii: list[dict[str, str]] = []
        itogo = None
        for row in self.ws.iter_rows(min_row=hdr.row + 1, max_row=hdr.row + 200):
            znacheniya = {c.column: _text(c) for c in row if _text(c)}
            pervaya = znacheniya.get(col_ids[0], "")
            if pervaya.isdigit():
                pozicii.append({
                    nazvanie: (format_chislo(znacheniya.get(col, "")) if col != col_ids[0] else znacheniya.get(col, ""))
                    for col, nazvanie in stolbcy
                })
            elif pozicii and znacheniya.get(col_ids[-1]):
                itogo = format_chislo(znacheniya[col_ids[-1]])
                break

        obekt_label = naiti_podpis(self.ws, "Основные характеристики объекта закупки")
        obekt_cell = sleduyuschaya_znachimaya_yacheyka(self.ws, obekt_label) if obekt_label else None

        self._polozhit(karta, "_nmck_tarif", hdr, {
            "obekt": _text(obekt_cell),
            "metod": _text(metod_cell),
            "reestr_tekst": stolbcy[1][1],
            "nds_zagolovok": stolbcy[2][1],
            "pozicii": pozicii,
            "itogo": itogo,
        })

    # ------------------------------------------------------------------ позиции
    def _pozicii(self, karta: Kartochka) -> None:
        zagolovok = naiti_podpis(self.ws, "Ед. изм.")
        if zagolovok is None:
            return
        stroka = zagolovok.row
        stolbcy: list[tuple[int, str]] = [(c.column, _text(c)) for c in self.ws[stroka] if _text(c)]
        if not any(n == "№" for _, n in stolbcy):
            for c in self.ws[stroka - 1]:
                if _text(c) == "№":
                    stolbcy.insert(0, (c.column, "№"))
                    break

        # объединённый заголовок (333зэа: «Источник финансирования» над J:K с
        # подзаголовками «софинансирование» / «грант») — читаются все его столбцы
        shirina = {}
        for col, _ in stolbcy:
            konec = col
            for rng in self.ws.merged_cells.ranges:
                if rng.min_row <= stroka <= rng.max_row and rng.min_col == col:
                    konec = rng.max_col
            while konec + 1 not in [c for c, _ in stolbcy] and not _text(self.ws.cell(stroka, konec + 1))                     and _text(self.ws.cell(stroka + 1, konec + 1)) and not _text(self.ws.cell(stroka + 1, 1)).isdigit():
                konec += 1
            shirina[col] = range(col, konec + 1)

        for row in self.ws.iter_rows(min_row=stroka + 1):
            nomer = _text(row[0])
            if not nomer.isdigit():
                continue
            poziciya: dict[str, Any] = {}
            for col, nazvanie in stolbcy:
                yachejki = [self.ws.cell(row[0].row, c) for c in shirina[col]]
                zapolnennye = [c for c in yachejki if _text(c)]
                cell = zapolnennye[0] if zapolnennye else yachejki[0]
                poziciya[nazvanie] = {
                    "znachenie": "; ".join(_text(c) for c in zapolnennye),
                    "fajl": self.fajl,
                    "list": self.list_,
                    "yacheyka": cell.coordinate,
                }
            karta.pozicii.append(poziciya)

    def izvlechi(
        self,
        podpisanty: dict[str, tuple[str, str]] | None = None,
        podrazdeleniya: dict[str, tuple[str, str]] | None = None,
        trebovaniya_standartnye: set[str] | None = None,
        sotrudniki: dict[str, str] | None = None,
        polzovatel_fio: str | None = None,
    ) -> Kartochka:
        karta = Kartochka()
        self._predmet(karta)
        self._ploshchadka_podskazka(karta)
        self._rekvizity_zayavki(karta)
        self._nds_tekst(karta)
        self._cena_vklyuchaet(karta)
        self._utverzhdenie(karta, podpisanty or {})
        self._soglasovanie(karta, podrazdeleniya or {})
        self._kontakt(karta, sotrudniki or {}, polzovatel_fio)
        self._usloviya_dogovora(karta)
        self._otv_dogovor(karta)
        self._instrukciya(karta)
        self._obespechenie(karta)
        self._nacrezhim(karta)
        self._trebovaniya_k_uchastnikam(karta, trebovaniya_standartnye or set())
        self._nmck_raschet(karta)
        self._nmck_tarif(karta)
        self._nmck_po_poziciyam(karta)
        self._pozicii(karta)
        # другая форма заявки (пример 426зэа): НДС и нацрежим — из строки
        # «Итого руб.: …» и из столбцов таблицы позиций
        if "nds_tekst" not in karta.znacheniya:
            self._nds_iz_itogo_rub(karta)
        if "nacrezhim_tekst" not in karta.znacheniya:
            self._nacrezhim_iz_pozicij(karta)
        return karta

    def _stolbec_pozicij(self, karta: Kartochka, *klyuchi: str) -> list[str]:
        znacheniya = []
        for p in karta.pozicii:
            for nazvanie, yach in p.items():
                if any(k in " ".join(nazvanie.lower().split()) for k in klyuchi):
                    znacheniya.append(str(yach.get("znachenie") or "").strip())
                    break
        return znacheniya

    def _nds_iz_itogo_rub(self, karta: Kartochka) -> None:
        """«Итого руб.: … в том числе НДС - 212661 (…) рубль 74 копейки» и
        столбец «Ставка НДС, %» → «в т.ч. НДС 22% – 212 661,74 руб.»."""
        cell = naiti_po_prefiksu(self.ws, "Итого руб.")
        tekst = _text(cell)
        m = re.search(r"в том числе НДС\s*[-–—]?\s*(\d[\d\s\xa0]*)\s*(?:\([^)]*\))?\s*руб\w*\s*(\d{1,2})\s*коп", tekst, re.I)
        if not m:
            if re.search(r"НДС не облагается|без НДС", tekst, re.I):
                self._polozhit(karta, "nds_tekst", cell, "НДС не облагается", tekst)
            return
        stavki = {s.replace(",", ".").rstrip("%").strip() for s in self._stolbec_pozicij(karta, "ставка ндс") if s}
        if len(stavki) != 1:
            return  # ставка не одна или не указана — пусть решит сотрудник
        stavka = stavki.pop()
        stavka = stavka[:-2] if stavka.endswith(".0") else stavka
        rubli = re.sub(r"[\s\xa0]", "", m.group(1))
        summa = format_chislo(Decimal(f"{rubli}.{int(m.group(2)):02d}"))
        self._polozhit(karta, "nds_tekst", cell, f"в т.ч. НДС {stavka}% – {summa} руб.", tekst)

    def _nacrezhim_iz_pozicij(self, karta: Kartochka) -> None:
        """Столбец «Применение Постановления №1875» таблицы позиций."""
        znacheniya = [z for z in self._stolbec_pozicij(karta, "1875") if z]
        if znacheniya:
            tekst = "; ".join(dict.fromkeys(znacheniya))
            self._polozhit(karta, "nacrezhim_tekst", None, tekst, tekst)


def izvlechi_kartochku(
    put_zayavki: str | Path,
    put_podpisantov: str | Path | None = None,
    put_podrazdelenij: str | Path | None = None,
    put_trebovanij_standartnyh: str | Path | None = None,
    put_sotrudnikov: str | Path | None = None,
    polzovatel_fio: str | None = None,
) -> Kartochka:
    """Извлечь карточку закупки из заявки Excel (раздел 9 ТЗ)."""
    podpisanty = zagruzit_podpisantov(put_podpisantov) if put_podpisantov else {}
    podrazdeleniya = zagruzit_podrazdeleniya(put_podrazdelenij) if put_podrazdelenij else {}
    standartnye = zagruzit_trebovaniya_standartnye(put_trebovanij_standartnyh) if put_trebovanij_standartnyh else set()
    sotrudniki = zagruzit_sotrudnikov(put_sotrudnikov) if put_sotrudnikov else {}
    return ZayavkaExcel(put_zayavki).izvlechi(podpisanty, podrazdeleniya, standartnye, sotrudniki, polzovatel_fio)
