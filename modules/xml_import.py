# -*- coding: utf-8 -*-
"""
Импорт данных из XML — дополнение к существующему импорту Excel.

Поддерживаемые форматы (определяются автоматически по корневому элементу):

  1. «ЕИС 223-ФЗ: извещение»  — выгрузка извещения с zakupki.gov.ru / площадки.
     Корень: <notice> или <export><notice>; ключевые поля — номер извещения,
     НМЦК, даты (подача, вскрытие, рассмотрение, итоги), предмет, лоты, ОКПД2.
  2. «ЕИС 223-ФЗ: протокол»   — <protocol>: участники, цены, отклонения, победитель.
  3. «Табличный XML»          — любой плоский XML вида <rows><row><col name="...">…
     или <rows><row><Поле>…</Поле></row></rows>; маппинг колонок — словарём
     COLUMN_MAP ниже или из базы знаний (статья «Табличный XML»).

Результат разбора — тот же dict, что даёт разбор docx/xlsx в панели
(поля: oao, pub, subject, nmck, dates, participants, winner, …),
поэтому дальше работает существующий submit_from_panel / handle_apply.

БЕЗОПАСНОСТЬ: документы с DOCTYPE, ENTITY, XInclude отклоняются до парсинга
(защита от XXE), размер файла ≤ 20 МБ (защита от «бомб» разрастания).
"""

from __future__ import annotations

import hashlib
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

MAX_SIZE = 20 * 1024 * 1024          # 20 МБ
MAX_ELEMENTS = 200_000               # защита от чрезмерно ветвистых документов

_FORBIDDEN = re.compile(rb'<!DOCTYPE|<!ENTITY|<xi:include', re.IGNORECASE)


class XmlImportError(ValueError):
    """Текст ошибки показывается пользователю в панели («на проверке»)."""


@dataclass
class ImportResult:
    format_key: str                    # 'eis223-notice' | 'eis223-protocol' | 'table-xml'
    item_kind: str                     # 'извещение' | 'протокол' | 'таблица'
    fields: dict = field(default_factory=dict)   # поля как у разбора docx/xlsx
    rows: list = field(default_factory=list)     # для table-xml — список строк
    warnings: list = field(default_factory=list)
    sha256: str = ''


# ------------------------------------------------------------
# Безопасный разбор
# ------------------------------------------------------------

def _safe_parse(source) -> ET.Element:
    raw = Path(source).read_bytes() if not isinstance(source, (bytes, bytearray)) else bytes(source)
    if len(raw) > MAX_SIZE:
        raise XmlImportError(f'Файл больше {MAX_SIZE // 1024 // 1024} МБ — не принимается')
    if _FORBIDDEN.search(raw[:8192]):
        raise XmlImportError('XML с DOCTYPE/ENTITY/XInclude не принимается (защита от XXE)')
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as e:
        raise XmlImportError(f'Файл не читается как XML: {e}')
    if sum(1 for _ in root.iter()) > MAX_ELEMENTS:
        raise XmlImportError('Слишком большой XML-документ')
    return root


def _local(tag: str) -> str:
    return tag.rsplit('}', 1)[-1]                       # срезаем namespace ЕИС


def _find_first(root, *names) -> ET.Element | None:
    wanted = {n.lower() for n in names}
    for el in root.iter():
        if _local(el.tag).lower() in wanted:
            return el
    return None


def _text(root, *names, default='') -> str:
    el = _find_first(root, *names)
    return (el.text or '').strip() if el is not None and el.text else default


def _norm_date(s: str) -> str:
    """'2026-09-03' / '03.09.2026' / '2026-09-03T12:00:00' → 'DD.MM.YYYY' (формат панели)."""
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', s or '')
    if m:
        return f'{m.group(3)}.{m.group(2)}.{m.group(1)}'
    m = re.match(r'(\d{2})\.(\d{2})\.(\d{4})', s or '')
    return m.group(0) if m else ''


def _money(s: str) -> str:
    s = re.sub(r'[^\d,.]', '', (s or '').replace('\u00a0', ''))
    return s.replace(',', '.')


# ------------------------------------------------------------
# Формат 1: извещение ЕИС / площадки (223-ФЗ)
# ------------------------------------------------------------

def _parse_notice(root) -> ImportResult:
    f = {
        'pub':      _text(root, 'noticeNumber', 'purchaseNumber', 'registrationNumber', 'number'),
        'subject':  _text(root, 'purchaseObjectInfo', 'subjectName', 'name', 'objectInfo'),
        'nmck':     _money(_text(root, 'maxPrice', 'initialSum', 'nmck')),
        'method':   _text(root, 'purchaseMethod', 'placingWayName', 'method'),
        'platform': _text(root, 'etpName', 'platform', 'etp'),
        'dates': {
            'publish':  _norm_date(_text(root, 'publishDate', 'publicationDate')),
            'open':     _norm_date(_text(root, 'openingDate', 'openDate', 'submissionCloseDate')),
            'review':   _norm_date(_text(root, 'examinationDate', 'reviewDate')),
            'results':  _norm_date(_text(root, 'summarizingDate', 'resultsDate')),
        },
        'contact_email': _text(root, 'email', 'contactEMail'),
        'contact_phone': _text(root, 'phone', 'contactPhone'),
        'okpd2': [_local_code(el) for el in root.iter() if _local(el.tag).lower() in ('okpd2', 'okpd2code')],
    }
    warnings = []
    if not f['subject']:
        warnings.append('Не найден предмет закупки (purchaseObjectInfo)')
    if not f['pub']:
        warnings.append('Не найден номер извещения')
    return ImportResult('eis223-notice', 'извещение', fields=f, warnings=warnings)


def _local_code(el) -> str:
    return (el.text or '').strip()


# ------------------------------------------------------------
# Формат 2: протокол ЕИС / площадки
# ------------------------------------------------------------

def _parse_protocol(root) -> ImportResult:
    f = {
        'pub':       _text(root, 'noticeNumber', 'purchaseNumber', 'registrationNumber'),
        'protocol_date': _norm_date(_text(root, 'protocolDate', 'signDate', 'createDate')),
        'participants': [],
    }
    for app in root.iter():
        if _local(app.tag).lower() not in ('application', 'participant', 'protocolapplication'):
            continue
        p = {
            'name':  _text(app, 'participantName', 'organizationName', 'name', 'fullName'),
            'inn':   _text(app, 'inn', 'participantInn'),
            'price': _money(_text(app, 'price', 'finalPrice', 'offerPrice')),
            'rejected': _text(app, 'rejected', 'isRejected').lower() in ('true', '1', 'да'),
            'reject_reason': _text(app, 'rejectReason', 'rejectionReason'),
            'winner': False,
        }
        if p['name'] or p['inn']:
            f['participants'].append(p)
    winner = _text(root, 'winnerName', 'winnerOrganization')
    winner_inn = _text(root, 'winnerInn')
    if f['participants']:
        if winner or winner_inn:
            for p in f['participants']:
                p['winner'] = (winner_inn and p['inn'] == winner_inn) or (winner and p['name'] == winner)
        if not any(p['winner'] for p in f['participants']):
            active = [p for p in f['participants'] if not p['rejected'] and p['price']]
            if active:                                   # как в панели: наименьшая цена
                min(active, key=lambda p: float(p['price'].replace(' ', '') or 'inf'))['winner'] = True
    f['winner'] = next((p['name'] for p in f['participants'] if p['winner']), '')
    f['winner_price'] = next((p['price'] for p in f['participants'] if p['winner']), '')
    warnings = [] if f['participants'] else ['Не найдены участники протокола']
    return ImportResult('eis223-protocol', 'протокол', fields=f, warnings=warnings)


# ------------------------------------------------------------
# Формат 3: табличный XML (универсальный)
# ------------------------------------------------------------

COLUMN_MAP = {                          # имя XML-поля → поле реестра (расширяется из базы знаний)
    'номер оао': 'oao', '№ оао': 'oao', 'oao': 'oao',
    'номер извещения': 'pub', 'pub': 'pub',
    'предмет': 'subject', 'предмет закупки': 'subject', 'subject': 'subject',
    'нмцк': 'nmck', 'nmck': 'nmck',
    'способ': 'method', 'method': 'method',
    'ответственный': 'responsible', 'responsible': 'responsible',
    'дата вскрытия': 'dt_open', 'дата итогов': 'dt_results',
    'победитель': 'winner', 'цена победителя': 'winner_price', 'инн': 'inn',
}


def _parse_table(root) -> ImportResult:
    found = _find_first(root, 'rows', 'table', 'data')
    rows_el = found if found is not None else root
    rows, warnings = [], []
    for row in rows_el:
        if _local(row.tag).lower() not in ('row', 'record', 'item', 'строка'):
            continue
        rec = {}
        for cell in row:
            name = (cell.get('name') or _local(cell.tag)).strip().lower()
            key = COLUMN_MAP.get(name, name)
            rec[key] = (cell.text or '').strip()
        if rec:
            rows.append(rec)
    if not rows:
        warnings.append('Табличных строк (<row>) не найдено')
    return ImportResult('table-xml', 'таблица', rows=rows, warnings=warnings)


# ------------------------------------------------------------
# Точка входа
# ------------------------------------------------------------

def detect_and_parse(source) -> ImportResult:
    """source — путь к файлу или bytes. → ImportResult (format_key, fields/rows, warnings)."""
    root = _safe_parse(source)
    kind = _local(root.tag).lower()
    raw = Path(source).read_bytes() if not isinstance(source, (bytes, bytearray)) else bytes(source)
    if kind in ('notice', 'purchase', 'epnotice') or _find_first(root, 'purchaseObjectInfo', 'noticeNumber') is not None \
            and _find_first(root, 'protocolDate') is None and 'protocol' not in kind:
        res = _parse_notice(root)
    elif 'protocol' in kind or _find_first(root, 'protocolDate', 'winnerName') is not None:
        res = _parse_protocol(root)
    else:
        res = _parse_table(root)
    res.sha256 = hashlib.sha256(raw).hexdigest()
    return res


def to_panel_payload(res: ImportResult) -> dict:
    """Вид для submit_from_panel/handle_apply панели."""
    if res.item_kind == 'таблица':
        return {'kind': 'table', 'rows': res.rows, 'format': res.format_key, 'sha256': res.sha256}
    return {'kind': res.item_kind, 'format': res.format_key, 'sha256': res.sha256,
            'warnings': res.warnings, **res.fields}


if __name__ == '__main__':
    import json, sys
    r = detect_and_parse(sys.argv[1])
    print(json.dumps(to_panel_payload(r), ensure_ascii=False, indent=2))
    if r.warnings:
        print('Предупреждения:', *r.warnings, sep='\n  • ', file=sys.stderr)
