# -*- coding: utf-8 -*-
"""
Вход по логину и паролю для панели УООЗД (хранение — MySQL).

Заменяет файловые «Пользователи.json» / «Панель — сессии.json».

Отличия от текущей файловой реализации:
  • PBKDF2-HMAC-SHA256, 210 000 итераций (OWASP 2023+), число итераций
    хранится в users.hash_iters и прозрачно повышается при входе;
  • блокировка учётной записи: 5 неудач → 15 минут;
  • таблица login_attempts — все попытки (и неудачные) с IP;
  • в БД лежит SHA-256 от токена сессии, сам токен — только в cookie;
  • сессию можно отозвать (кнопка «Завершить сеансы» у администратора);
  • парольная политика: ≥ 10 знаков, ≠ логину, не из списка типовых.

Интеграция в reestr_server.py (точки замены):
    check_login(login, password)   →  Auth(db).login(login, password, ip)
    new_session(login)             →  Auth(db).open_session(user_id, ip, ua)
    user_by_session(token)         →  Auth(db).user_by_token(token)
    drop_session(token)            →  Auth(db).close_session(token)
    set_password(u, pw)            →  Auth(db).set_password(user_id, pw)
"""

from __future__ import annotations

import hashlib
import hmac
import re
import secrets
import time
from datetime import datetime, timedelta

HASH_ITERS = 210_000
SESSION_HOURS = 12            # боевой режим; в файловом было 30 суток
LOCK_AFTER = 5                # неудачных попыток подряд
LOCK_MINUTES = 15

COMMON_PASSWORDS = {
    '123456', '12345678', '1234567890', 'qwerty', 'qwerty123', 'password',
    'password1', '11111111', '00000000', 'admin123', 'sechenov', 'uoozd123',
    'йцукен', 'йцукенгш', 'пароль', 'пароль123',
}


class AuthError(ValueError):
    pass


def _pbkdf2(password: str, salt_hex: str, iters: int) -> str:
    return hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'),
                               bytes.fromhex(salt_hex), iters).hex()


def password_policy(login: str, password: str) -> None:
    """Правила пароля. Нарушение → AuthError с текстом для пользователя."""
    if len(password) < 10:
        raise AuthError('Пароль — не короче 10 знаков')
    if password.lower() == login.lower():
        raise AuthError('Пароль не должен совпадать с логином')
    if password.lower() in COMMON_PASSWORDS:
        raise AuthError('Слишком простой пароль — он есть в списке типовых')
    if not re.search(r'[A-Za-zА-Яа-яЁё]', password) or not re.search(r'\d', password):
        raise AuthError('Пароль должен содержать буквы и цифры')


class Auth:
    def __init__(self, db):
        self.db = db          # modules.db.DB

    # ---------- пароль ----------

    def set_password(self, user_id: int, password: str, *, must_change=False):
        user = self.db.query_one('SELECT login FROM users WHERE id=%s', (user_id,))
        if not user:
            raise AuthError('Пользователь не найден')
        password_policy(user['login'], password)
        salt = secrets.token_hex(16)
        self.db.execute(
            'UPDATE users SET salt=%s, pass_hash=%s, hash_iters=%s, must_change=%s, '
            '                 pwd_changed_at=NOW(), failed_logins=0, locked_until=NULL WHERE id=%s',
            (salt, _pbkdf2(password, salt, HASH_ITERS), HASH_ITERS, int(must_change), user_id))

    def _verify(self, u, password) -> bool:
        ok = hmac.compare_digest(u['hash'], _pbkdf2(password, u['salt'], u['hash_iters']))
        if ok and u['hash_iters'] < HASH_ITERS:
            # прозрачное повышение стойкости при следующем входе
            self.set_password(u['id'], password)
        return ok

    # ---------- вход ----------

    def login(self, login: str, password: str, ip: bytes | None = None):
        """→ dict пользователя. Ошибка → AuthError с безопасным текстом (без «логин существует, а пароль нет»)."""
        login = (login or '').strip().lower()
        u = self.db.query_one(
            'SELECT id, login, fio, dept, is_admin AS admin, can_delete, must_change, is_active, '
            '       salt, pass_hash AS hash, hash_iters, failed_logins, locked_until '
            'FROM users WHERE login=%s', (login,))

        if u and u['locked_until'] and u['locked_until'] > datetime.now():
            self._attempt(login, ip, ok=False)
            wait = int((u['locked_until'] - datetime.now()).total_seconds() // 60) + 1
            raise AuthError(f'Учётная запись заблокирована. Повторите через {wait} мин.')

        ok = bool(u) and u['is_active'] and self._verify(u, password or '')
        self._attempt(login, ip, ok)

        with self.db.cursor(commit=True) as cur:
            if ok:
                cur.execute('UPDATE users SET failed_logins=0, locked_until=NULL WHERE id=%s', (u['id'],))
            elif u:
                fails = u['failed_logins'] + 1
                lock = f', locked_until = DATE_ADD(NOW(), INTERVAL {LOCK_MINUTES} MINUTE)' if fails >= LOCK_AFTER else ''
                cur.execute(f'UPDATE users SET failed_logins=%s{lock} WHERE id=%s', (fails, u['id']))
        if not ok:
            time.sleep(0.8)    # ровная пауза: по времени ответа нельзя отличить «нет логина» от «неверный пароль»
            raise AuthError('Неверный логин или пароль')
        return {k: u[k] for k in ('id', 'login', 'fio', 'dept', 'admin', 'can_delete', 'must_change')}

    def _attempt(self, login, ip, ok):
        self.db.execute('INSERT INTO login_attempts (login, ip, ok) VALUES (%s,%s,%s)',
                        (login[:32], ip, int(ok)))

    # ---------- сессии ----------

    @staticmethod
    def _token_hash(token: str) -> str:
        return hashlib.sha256(token.encode('utf-8')).hexdigest()

    def open_session(self, user_id: int, ip: bytes | None = None, user_agent: str = '') -> str:
        token = secrets.token_urlsafe(32)
        self.db.execute(
            'INSERT INTO sessions (token_hash, user_id, ip, user_agent, expires_at) '
            'VALUES (%s,%s,%s,%s, DATE_ADD(NOW(), INTERVAL %s HOUR))',
            (self._token_hash(token), user_id, ip, (user_agent or '')[:255], SESSION_HOURS))
        return token

    def user_by_token(self, token: str):
        if not token:
            return None
        return self.db.query_one(
            'SELECT u.id, u.login, u.fio, u.dept, u.is_admin AS admin, u.can_delete, u.must_change '
            'FROM sessions s JOIN users u ON u.id = s.user_id '
            'WHERE s.token_hash=%s AND s.revoked=0 AND s.expires_at > NOW() AND u.is_active=1',
            (self._token_hash(token),))

    def close_session(self, token: str):
        if token:
            self.db.execute('UPDATE sessions SET revoked=1 WHERE token_hash=%s', (self._token_hash(token),))

    def revoke_all(self, user_id: int):
        self.db.execute('UPDATE sessions SET revoked=1 WHERE user_id=%s', (user_id,))

    def purge_expired(self):
        self.db.execute('DELETE FROM sessions WHERE expires_at < DATE_SUB(NOW(), INTERVAL 7 DAY)')


# ---------- первый администратор (как ensure_first_admin сейчас) ----------

def ensure_first_admin(db, login='admin', fio='Администратор'):
    """Пользователей нет → создаём администратора и возвращаем (login, пароль) для печати в консоль."""
    if db.query_one('SELECT id FROM users LIMIT 1'):
        return None
    password = secrets.token_urlsafe(9)   # ~12 символов
    salt = secrets.token_hex(16)
    db.execute(
        'INSERT INTO users (login, fio, dept, is_admin, can_delete, must_change, salt, pass_hash, hash_iters) '
        'VALUES (%s,%s,\'ОАО\',1,1,1,%s,%s,%s)',
        (login, fio, salt, _pbkdf2(password, salt, HASH_ITERS), HASH_ITERS))
    return login, password
