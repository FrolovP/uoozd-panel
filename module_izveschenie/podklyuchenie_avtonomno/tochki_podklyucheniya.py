# -*- coding: utf-8 -*-
"""Пять точек подключения (раздел 14 ТЗ) — автономная версия.

Ядро (dokumentaciya/) получает всё внешнее только через эти пять функций
(объект Nastrojki.tochki). В автономной версии точки 2–5 ничего не делают
и возвращают VYKLYUCHENO — тогда ядро работает само: читает заявку по
правилам, PDF не читает, ИИ не вызывает. Точка 1 «Настройки» обязательна:
здесь она берёт значения из config.ini.

При встраивании этот файл заменяется своим — с теми же именами функций и
тем же смыслом ответов; код ядра не меняется (VSTRAIVANIE.md).
"""
from __future__ import annotations

import configparser
from pathlib import Path
from typing import Any

VYKLYUCHENO = "выключено"
# ответ точки 2–5, если она подключена, но её сервер сейчас не отвечает (можно
# и просто выбросить исключение) — модуль не падает, работает сам и пишет
# сотруднику «Сервер недоступен, попробуйте позже»
NEDOSTUPEN = "недоступен"

KOREN = Path(__file__).resolve().parents[1]


def nastrojki() -> dict[str, Any]:
    """1. Настройки: путь к данным, общая папка справочников, папка карточек
    закупок, пользователь, порт и версия. Автономно — из config.ini;
    put_obshchih_spravochnikov None — справочники из <путь к данным>/spravochniki;
    put_zakupok None — карточки в <путь к данным>/zakupki; пользователь
    None — сотрудник выбирает себя на экране «Кто вы»."""
    cfg = configparser.ConfigParser()
    cfg.read(KOREN / "config.ini", encoding="utf-8")
    obshchie = cfg.get("programma", "папка_справочников", fallback="").strip()
    obshchie_zakupki = cfg.get("programma", "папка_закупок", fallback="").strip()
    return {
        "put_dannyh": KOREN / cfg.get("programma", "put_dannyh", fallback="dannye"),
        "put_obshchih_spravochnikov": (KOREN / obshchie) if obshchie else None,
        "put_obshchih_zakupok": (KOREN / obshchie_zakupki) if obshchie_zakupki else None,
        "put_zakupok": None,
        "polzovatel": None,
        "port": cfg.getint("programma", "port", fallback=8765),
        "versiya": cfg.get("programma", "versiya", fallback=""),
    }


def raspoznannyj_tekst(put_fajla: str | Path) -> str:
    """2. Распознанный текст PDF или скана (OCR-сервер программы): текст
    документа. Автономно — VYKLYUCHENO (PDF получает роль «Не используется»)."""
    return VYKLYUCHENO


def dannye_proverki_zayavki(put_zayavki: str | Path) -> dict[str, Any] | str:
    """3. Данные проверки заявки из вкладки программы: {имя поля карточки:
    значение} — берутся первыми, остальное ядро читает из файла заявки.
    Автономно — VYKLYUCHENO (всё читается из файла)."""
    return VYKLYUCHENO


def sopostavlenie_stolbcov(zagolovki: list[str]) -> dict[str, str] | str:
    """4. Сопоставление столбцов таблицы позиций: {заголовок в заявке:
    заголовок для карточки}. Автономно — VYKLYUCHENO (заголовки как в заявке,
    своё сопоставление ядра по подписям)."""
    return VYKLYUCHENO


def vyzov_ii(vopros: str) -> str:
    """5. Вызов ИИ: «вопрос → ответ» поверх слоя выбора провайдера программы.
    Автономно — VYKLYUCHENO (модуль работает без ИИ)."""
    return VYKLYUCHENO
