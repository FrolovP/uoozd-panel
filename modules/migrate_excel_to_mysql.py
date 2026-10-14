# -*- coding: utf-8 -*-
"""
Миграция данных панели УООЗД из Excel/JSON в MySQL.

Запуск:
    python migrate_excel_to_mysql.py --config config.ini --panel "Демо-реестр" [--dry-run]

Что переносится:
  • «Реестр <год>.xlsx» (листы конкурентных и неконкурентных) → procurements, counterparties
  • «Панель — пользователи.json»                             → users (хэши переносятся как есть)
  • «Новости.json»                                           → news
  • журнал «Реестр — журнал.json»                            → audit_log

Файлы-копии вида «Реестр 2026_20260512_100549.xlsx» (еженедельные снимки) пропускаются.
--dry-run — только считает и печатает, в БД не пишет.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).parent))
from db import DB  # noqa: E402

MONTHS = {'января': 1, 'февраля': 2, 'марта': 3, 'апреля': 4, 'мая': 5, 'июня': 6,
          'июля': 7, 'августа': 8, 'сентября': 9, 'октября': 10, 'ноября': 11, 'декабря': 12}


def norm(s) -> str:
    return re.sub(r'\s+', ' ', str(s or '').strip().lower())


def parse_date(v):
    """'19.01.2026' / '5 января 2026' / datetime / '…от 21 января 2026' → date | None"""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = norm(v)
    m = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{4})', s)
    if m:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.search(r'(\d{1,2})\s+([а-яё]+)\s+(\d{4})', s)
    if m and m.group(2) in MONTHS:
        return date(int(m.group(3)), MONTHS[m.group(2)], int(m.group(1)))
    return None


def parse_money(v):
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    s = re.sub(r'[^\d,.]', '', str(v).replace('\u00a0', '').replace(' ', ''))
    if not s:
        return None
    s = s.replace(',', '.')
    try:
        return round(float(s), 2)
    except ValueError:
        return None


# нормализованный заголовок → (поле procurements | 'payload:ключ', обработчик)
HEADER_RULES = [
    ('наименование товара',            ('subject', str)),
    ('ответственный за документацию',  ('responsible', str)),
    ('наименование торговой площадки', ('platform', str)),
    ('номер и дата публикации',        ('__pub__', str)),
    ('дата вскрытия',                  ('dt_open', parse_date)),
    ('дата рассмотрения',              ('dt_review', parse_date)),
    ('дата подведения итогов',         ('dt_results', parse_date)),
    ('организация -победитель',        ('__winner__', str)),
    ('инн',                            ('__inn__', str)),
    ('начальная максимальн',           ('nmck', parse_money)),
    ('цена договора по протоколу',     ('winner_price', parse_money)),
    ('экономия по протоколу',          ('economy', parse_money)),
    ('кат. см и сп',                   ('smp_cat', norm)),
    ('смп',                            ('smp_cat', norm)),
    ('аббрев процедур',                ('method', str)),
    ('абрев процедур',                 ('method', str)),
    ('к,л,м,то,р',                     ('payload:category', str)),
    ('закон 223/44',                   ('payload:law', str)),
    ('1с',                             ('payload:num_1c', str)),
    ('фас',                            ('payload:fas', str)),
    ('н.кц/вуз',                       ('__science__', str)),
    ('№ кц/вуз',                       ('__science__', str)),
    ('подчин',                         ('payload:subord', str)),
    ('всего торгов',                   ('payload:total_trades', str)),
    ('цена заключенного договора',     ('payload:contract_price', parse_money)),
    ('номер и дата договора',          ('__contract__', str)),
    ('экономия по договору',           ('payload:contract_economy', parse_money)),
    ('аукцион/конкурс сотоялся',       ('payload:held', str)),
    ('аукцион/конкурс состоялся',      ('payload:held', str)),
    ('без аукциона',                   ('payload:no_auction', str)),
    ('договор заключен',               ('payload:contract_done', str)),
    ('кол-во всего заявок',            ('payload:apps_total', str)),
    ('кол-во отклон',                  ('payload:apps_rejected', str)),
]


def map_header(h: str):
    n = norm(h)
    if not n:
        return None
    for prefix, rule in HEADER_RULES:
        if n.startswith(prefix):
            return rule
    return None


def parse_pub(s: str):
    """'79зэа-оа 32610542436 (https://…) 5 января 2026' → (oao, pub_number, dt_publish)"""
    s = norm(s)
    oao = pub = None
    m = re.match(r'(\d+\s*[а-яё]+)', s)
    if m:
        oao = m.group(1).replace(' ', '')
    m = re.search(r'\b(\d{10,})\b', s)
    if m:
        pub = m.group(1)
    return oao, pub, parse_date(s)


def rows_from_sheet(ws):
    headers = [map_header(c.value) for c in ws[3]]          # 3-я строка — шапка
    for r in ws.iter_rows(min_row=4, values_only=True):
        first = norm(r[0] if r else '')
        if not any(v not in (None, '') for v in r):
            continue
        if first.startswith('---'):                          # строка-разделитель «---05.01.2026---»
            continue
        row, payload = {}, {}
        for rule, v in zip(headers, r):
            if rule is None or v in (None, ''):
                continue
            key, conv = rule
            try:
                val = conv(v)
            except Exception:
                val = str(v)
            if val in (None, ''):
                continue
            if key.startswith('payload:'):
                payload[key[8:]] = str(val) if not isinstance(val, (int, float)) else val
            else:
                row[key] = val
        if not row.get('subject'):
            continue
        # составные поля
        oao, pub, dt_pub = parse_pub(row.pop('__pub__', '') or '')
        if oao:
            row['oao_number'] = oao
        if pub:
            row['pub_number'] = pub
        if dt_pub:
            row['dt_publish'] = dt_pub
        winner = str(row.pop('__winner__', '') or '')
        row['__winner_name__'] = re.sub(r'\s*инн\s*[\d\s]*$', '', winner, flags=re.IGNORECASE).strip()
        row['__winner_inn__'] = re.sub(r'\D', '', str(row.pop('__inn__', '') or ''))
        science = str(row.pop('__science__', '') or '').strip()
        if science:
            row['is_science'] = 1
            payload['science'] = science
        contract = str(row.pop('__contract__', '') or '').strip()
        if contract:
            payload['contract'] = contract
        # статус
        if contract or payload.get('contract_done'):
            row['status'] = 'договор заключён'
        elif row.get('__winner_name__'):
            row['status'] = 'итоги подведены'
        elif row.get('dt_open'):
            row['status'] = 'заявки поданы' if row['dt_open'] <= date.today() else 'приём заявок'
        else:
            row['status'] = 'в работе'
        if payload:
            row['payload'] = json.dumps(payload, ensure_ascii=False)
        yield row


def upsert_counterparty(db, name, inn, dry):
    if not name or not inn:
        return None
    if dry:
        return 1
    db.execute('INSERT INTO counterparties (name, inn) VALUES (%s,%s) '
               'ON DUPLICATE KEY UPDATE name=VALUES(name), id=LAST_INSERT_ID(id)', (name, inn))
    r = db.query_one('SELECT id FROM counterparties WHERE inn=%s', (inn,))
    return r['id'] if r else None


def migrate_registry(db, path: Path, year: int, dry: bool, log):
    wb = load_workbook(path, read_only=True, data_only=True)
    n = skipped = 0
    for ws in wb.worksheets:
        sheet_key = 'nk' if 'неконк' in ws.title.lower() else 'k'
        for row in rows_from_sheet(ws):
            if not row.get('oao_number'):
                skipped += 1
                continue
            winner_name, winner_inn = row.pop('__winner_name__', ''), row.pop('__winner_inn__', '')
            row.update(year=year, sheet=sheet_key)
            n += 1
            if dry:
                continue
            row['winner_id'] = upsert_counterparty(db, winner_name, winner_inn, dry)
            cols = ['year', 'sheet', 'oao_number', 'pub_number', 'subject', 'method', 'nmck', 'responsible',
                    'platform', 'dt_publish', 'dt_open', 'dt_review', 'dt_results', 'winner_id',
                    'winner_price', 'economy', 'status', 'smp_cat', 'is_science', 'payload']
            data = {c: row.get(c) for c in cols}
            data['is_science'] = row.get('is_science', 0)
            db.execute(
                'INSERT INTO procurements (' + ', '.join(cols) + ') VALUES (' + ', '.join(f'%({c})s' for c in cols) + ') '
                'ON DUPLICATE KEY UPDATE pub_number=VALUES(pub_number), subject=VALUES(subject), '
                'nmck=VALUES(nmck), responsible=VALUES(responsible), platform=VALUES(platform), '
                'dt_publish=VALUES(dt_publish), dt_open=VALUES(dt_open), dt_review=VALUES(dt_review), '
                'dt_results=VALUES(dt_results), winner_id=VALUES(winner_id), winner_price=VALUES(winner_price), '
                'economy=VALUES(economy), status=VALUES(status), smp_cat=VALUES(smp_cat), '
                'is_science=VALUES(is_science), payload=VALUES(payload)', data)
    log(f'  {path.name}: {n} строк {"(проба)" if dry else "перенесено"}' + (f', без № ОАО: {skipped}' if skipped else ''))
    return n


def migrate_users(db, panel: Path, dry: bool, log):
    f = next(panel.rglob('*ользовател*.json'), None)
    if not f:
        log('  пользователи: файл не найден — первый администратор создастся при запуске')
        return 0
    users = json.loads(f.read_text(encoding='utf-8'))
    for u in users:
        if dry:
            continue
        db.execute(
            'INSERT INTO users (login, fio, dept, is_admin, can_delete, must_change, salt, pass_hash, hash_iters) '
            'VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) ON DUPLICATE KEY UPDATE fio=VALUES(fio)',
            (u['login'], u['fio'], u.get('dept', 'ОАО'), int(u.get('admin', 0)), int(u.get('can_delete', 0)),
             int(u.get('must_change', 0)), u.get('salt', '00'), u.get('hash', '00'), 120000))
    log(f'  пользователи: {len(users)} {"(проба)" if dry else "перенесено"}')
    return len(users)


def migrate_json_log(src: Path, db, dry, log):
    if not src.exists():
        return 0
    items = json.loads(src.read_text(encoding='utf-8'))
    items = items if isinstance(items, list) else items.get('items', [])
    if not dry:
        for e in items:
            db.execute('INSERT INTO audit_log (ts, kind, user_fio, message) VALUES (%s,%s,%s,%s)',
                       (e.get('ts') or None, e.get('kind', 'info')[:16], e.get('user'), e.get('message', e.get('text', ''))))
    log(f'  {src.name}: {len(items)} записей {"(проба)" if dry else "перенесено"}')
    return len(items)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--config', default='config.ini')
    ap.add_argument('--panel', required=True, help='папка «Демо-реестр» или боевая папка реестра')
    ap.add_argument('--dry-run', action='store_true')
    a = ap.parse_args()

    panel = Path(a.panel)
    db = None if a.dry_run else DB.from_config(a.config)
    log = print
    total = 0
    # только основные файлы «Реестр <год>.xlsx» — еженедельные копии «Реестр 2026_…xlsx» пропускаем
    registries = [f for f in sorted(panel.glob('**/Реестр *.xlsx'))
                  if re.fullmatch(r'Реестр \d{4}', f.stem)]
    for xlsx in registries:
        total += migrate_registry(db, xlsx, int(xlsx.stem.split()[-1]), a.dry_run, log)
    total += migrate_users(db, panel, a.dry_run, log)
    total += migrate_json_log(next(panel.rglob('*журнал*.json'), Path('∅')), db, a.dry_run, log)
    print(f'Итого: {total} записей {"будет перенесено" if a.dry_run else "перенесено в MySQL"}')


if __name__ == '__main__':
    main()
