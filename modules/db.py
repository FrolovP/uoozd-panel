# -*- coding: utf-8 -*-
"""
Слой доступа к MySQL для панели УООЗД.

Зависимости:  pip install pymysql cryptography
Подключение настраивается в config.ini (см. ARCHITECTURE.md, раздел 5),
пароль — только из переменной окружения UOOZD_DB_PASSWORD.

Использование:
    from db import DB
    db = DB.from_config('config.ini')
    with db.cursor() as cur:
        cur.execute('SELECT ... WHERE year = %s', (2026,))
        rows = cur.fetchall()

Только параметризованные запросы (%s) — конкатенация строк в SQL запрещена.
"""

from __future__ import annotations

import configparser
import os
import queue
import threading
from contextlib import contextmanager

import pymysql
import pymysql.cursors


class DBError(RuntimeError):
    pass


class DB:
    """Пул соединений MySQL (потокобезопасный)."""

    def __init__(self, *, host, port, database, user, password, ssl_ca=None, pool_size=8, connect_timeout=5):
        self._args = dict(
            host=host, port=port, db=database, user=user, password=password,
            charset='utf8mb4', cursorclass=pymysql.cursors.DictCursor,
            autocommit=False, connect_timeout=connect_timeout,
            read_timeout=30, write_timeout=30,
        )
        if ssl_ca:
            self._args['ssl'] = {'ca': ssl_ca}
        self._pool: queue.Queue = queue.Queue(maxsize=pool_size)
        self._lock = threading.Lock()
        self._created = 0
        self._pool_size = pool_size

    # ---------- создание ----------

    @classmethod
    def from_config(cls, path='config.ini') -> 'DB':
        cp = configparser.ConfigParser()
        if not cp.read(path, encoding='utf-8'):
            raise DBError(f'Не найден конфиг: {path}')
        m = cp['mysql']
        password = os.environ.get('UOOZD_DB_PASSWORD', '')
        if not password:
            raise DBError('Задайте пароль БД в переменной окружения UOOZD_DB_PASSWORD')
        return cls(host=m.get('host', '127.0.0.1'), port=m.getint('port', 3306),
                   database=m.get('database', 'uoozd'), user=m.get('user', 'uoozd_app'),
                   password=password, ssl_ca=m.get('ssl_ca', fallback=None) or None,
                   pool_size=m.getint('pool_size', 8))

    def _connect(self):
        try:
            return pymysql.connect(**self._args)
        except pymysql.MySQLError as e:
            raise DBError(f'Нет соединения с MySQL: {e.args[1] if e.args else e}') from e

    def _take(self):
        try:
            conn = self._pool.get_nowait()
            try:
                conn.ping(reconnect=True)
            except pymysql.MySQLError:
                return self._connect()
            return conn
        except queue.Empty:
            with self._lock:
                if self._created < self._pool_size:
                    self._created += 1
                    return self._connect()
            return self._pool.get()

    def _put(self, conn):
        try:
            self._pool.put_nowait(conn)
        except queue.Full:
            conn.close()
            with self._lock:
                self._created -= 1

    # ---------- работа ----------

    @contextmanager
    def cursor(self, *, commit=False):
        """Соединение из пула + курсор. commit=True — транзакция фиксируется, при ошибке откатывается."""
        conn = self._take()
        try:
            with conn.cursor() as cur:
                yield cur
            if commit:
                conn.commit()
            else:
                conn.rollback()
        except Exception:
            conn.rollback()
            raise
        finally:
            self._put(conn)

    def query(self, sql, params=()):
        with self.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchall()

    def query_one(self, sql, params=()):
        rows = self.query(sql, params)
        return rows[0] if rows else None

    def execute(self, sql, params=()):
        with self.cursor(commit=True) as cur:
            cur.execute(sql, params)
            return cur.rowcount

    def close(self):
        while True:
            try:
                self._pool.get_nowait().close()
            except queue.Empty:
                break


# ============================================================
# Репозитории: те же ответы, что сейчас дают load_workbook/_json_file,
# чтобы сервисный слой панели менялся только в точках вызова.
# ============================================================

class UsersRepo:
    def __init__(self, db: DB):
        self.db = db

    def all(self):
        return self.db.query(
            'SELECT id, login, fio, dept, is_admin AS admin, can_delete, must_change, is_active, '
            '       salt, pass_hash AS hash, hash_iters, failed_logins, locked_until '
            'FROM users ORDER BY fio')

    def by_login(self, login):
        return self.db.query_one(
            'SELECT id, login, fio, dept, is_admin AS admin, can_delete, must_change, is_active, '
            '       salt, pass_hash AS hash, hash_iters, failed_logins, locked_until '
            'FROM users WHERE login = %s', (login.lower().strip(),))

    def save(self, u: dict):
        with self.db.cursor(commit=True) as cur:
            if u.get('id'):
                cur.execute(
                    'UPDATE users SET fio=%s, dept=%s, is_admin=%s, can_delete=%s, must_change=%s, '
                    '                 is_active=%s, salt=%s, pass_hash=%s, hash_iters=%s WHERE id=%s',
                    (u['fio'], u['dept'], u['admin'], u['can_delete'], u['must_change'],
                     u.get('is_active', 1), u['salt'], u['hash'], u.get('hash_iters', 210000), u['id']))
            else:
                cur.execute(
                    'INSERT INTO users (login, fio, dept, is_admin, can_delete, must_change, salt, pass_hash, hash_iters) '
                    'VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)',
                    (u['login'], u['fio'], u['dept'], u['admin'], u['can_delete'], u['must_change'],
                     u['salt'], u['hash'], u.get('hash_iters', 210000)))
                u['id'] = cur.lastrowid
        return u


class RegistryRepo:
    """Реестр ОАО: замена чтения «Реестр <год>.xlsx»."""

    COLUMNS = ('id, year, sheet, oao_number, pub_number, subject, method, nmck, responsible, platform, '
               'dt_publish, dt_open, dt_review, dt_results, dt_contract_plan, winner_price, economy, '
               'status, smp_cat, is_science, deleted, payload, updated_at')

    def __init__(self, db: DB):
        self.db = db

    def by_year(self, year, sheet=None, include_deleted=False):
        sql = f'SELECT {self.COLUMNS} FROM procurements WHERE year = %s'
        params = [year]
        if sheet:
            sql += ' AND sheet = %s'
            params.append(sheet)
        if not include_deleted:
            sql += ' AND deleted = 0'
        return self.db.query(sql + ' ORDER BY dt_open IS NULL, dt_open, id', params)

    def find(self, year, oao_number):
        return self.db.query_one(
            f'SELECT {self.COLUMNS} FROM procurements WHERE year = %s AND oao_number = %s AND deleted = 0',
            (year, oao_number))

    def upsert(self, row: dict, author=''):
        keys = [k for k in row if k != 'id']
        with self.db.cursor(commit=True) as cur:
            if row.get('id'):
                cur.execute(
                    'UPDATE procurements SET ' + ', '.join(f'{k}=%s' for k in keys) + ' WHERE id=%s',
                    [row[k] for k in keys] + [row['id']])
                return row['id']
            cur.execute(
                'INSERT INTO procurements (' + ', '.join(keys) + ') VALUES (' +
                ', '.join(['%s'] * len(keys)) + ') '
                'ON DUPLICATE KEY UPDATE id = LAST_INSERT_ID(id), ' +
                ', '.join(f'{k}=VALUES({k})' for k in keys if k not in ('year', 'sheet', 'oao_number')),
                [row[k] for k in keys])
            return cur.lastrowid

    def max_updated(self, year):
        r = self.db.query_one('SELECT MAX(updated_at) AS m, COUNT(*) AS n FROM procurements WHERE year=%s', (year,))
        return (r['m'], r['n']) if r else (None, 0)   # → ETag для /api/data


class JournalRepo:
    """Журнал — только добавление."""

    def __init__(self, db: DB):
        self.db = db

    def add(self, kind, message, user_fio=None, ip=None):
        self.db.execute('INSERT INTO audit_log (kind, user_fio, ip, message) VALUES (%s,%s,%s,%s)',
                        (kind, user_fio, ip, message))

    def since(self, last_id=0, limit=500):
        return self.db.query('SELECT id, ts, kind, user_fio, message FROM audit_log '
                             'WHERE id > %s ORDER BY id LIMIT %s', (last_id, limit))
